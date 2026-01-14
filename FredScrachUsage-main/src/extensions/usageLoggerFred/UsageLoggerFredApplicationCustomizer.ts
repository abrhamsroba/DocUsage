import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
 
export interface IUsageLoggerFredApplicationCustomizerProperties {
  LoggingSiteUrl?: string; // optional: central logging site
  ListTitle: string;       // e.g., "Site Usage Data"
  ThrottleMs?: number;     // optional delay before logging
}
 
const LOG_SOURCE: string = 'UsageLoggerFredApplicationCustomizer';
 
export default class UsageLoggerFredApplicationCustomizer
  extends BaseApplicationCustomizer<IUsageLoggerFredApplicationCustomizerProperties> {
 
  public async onInit(): Promise<void> {
    try {
      // prevent duplicate log on same page load
      const pageKey = `usageLogged:${location.href}`;
      if (sessionStorage.getItem(pageKey)) return;
 
      const delay = this.properties?.ThrottleMs ?? 0;
      if (delay > 0) await new Promise(r => setTimeout(r, delay));
 
      const pageUrl = location.href.split('?')[0];
      const pageTitle = (document.title || '').substring(0, 255);
      const referrer = (document.referrer || '').substring(0, 255);
      const siteUrl = this.context.pageContext.site.absoluteUrl;
      const webUrl  = this.context.pageContext.web.absoluteUrl;
      const userDisp = (this.context.pageContext.user.displayName || '').substring(0, 255);
      const clientInfo = this._getClientInfo().substring(0, 255);
      const sessionId = this._ensureSessionId();
      const isHome = this._isHomePage();
 
      const targetSite = this.properties?.LoggingSiteUrl || webUrl;
      const listTitle  = this.properties?.ListTitle || 'Site Usage Data';
 
      await this._addItem(targetSite, listTitle, {
        PageUrl: pageUrl,
        PageTitle: pageTitle,
        Referrer: referrer,
        SiteUrl: siteUrl,
        WebUrl: webUrl,
        UserDisplayName: userDisp,
        SessionId: sessionId,
        IsHomePage: isHome,
        ClientInfo: clientInfo,
        TimeStamp: new Date().toISOString()
      });
 
      sessionStorage.setItem(pageKey, '1');
      Log.info(LOG_SOURCE, 'Usage item created.');
    } catch (e) {
      Log.warn(LOG_SOURCE, e as any);
    }
  }
 
  private _getClientInfo(): string {
    const w: any = window, s = screen;
    const dpr = w.devicePixelRatio || 1;
    return `ua:${navigator.userAgent}|wh:${w.innerWidth}x${w.innerHeight}|scr:${s?.width}x${s?.height}|dpr:${dpr}`;
  }
 
  private _ensureSessionId(): string {
    let id = sessionStorage.getItem('usageSessionId');
    if (!id) {
      id = Math.random().toString(36).slice(2) + Date.now().toString(36);
      sessionStorage.setItem('usageSessionId', id);
    }
    return id;
  }
 
  private _isHomePage(): boolean {
    try {
      const serverRel = this.context.pageContext.site.serverRelativeUrl?.replace(/\/$/, '') || '';
      const current = (this.context.pageContext.site as any).serverRequestPath?.replace(/\/$/, '')
        || location.pathname.replace(/\/$/, '');
      return current === '' || current === '/' || current === serverRel;
    } catch { return false; }
  }
 
  private async _addItem(targetSiteUrl: string, listTitle: string, payload: any): Promise<void> {
    const endpoint = `${targetSiteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items`;
    const resp: SPHttpClientResponse = await this.context.spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata'
        },
        body: JSON.stringify(payload)
      }
    );
    if (!resp.ok) {
      const text = await resp.text();
      throw new Error(`Usage log failed: ${resp.status} ${resp.statusText} - ${text}`);
    }
  }
}