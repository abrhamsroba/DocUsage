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

      // ============================================================
      // MICROSOFT CLARITY
      // ============================================================

      const clarityScript = document.createElement("script");

      clarityScript.type = "text/javascript";

      clarityScript.text = `
        (function(c,l,a,r,i,t,y){
          c[a]=c[a]||function(){
            (c[a].q=c[a].q||[]).push(arguments)
          };
          t=l.createElement(r);
          t.async=1;
          t.src="https://www.clarity.ms/tag/"+i;
          y=l.getElementsByTagName(r)[0];
          y.parentNode.insertBefore(t,y);
        })(window, document, "clarity", "script", "v0ywy6glon");
      `;

      document.head.appendChild(clarityScript);


      // ============================================================
      // START DOCUMENT TRACKING
      // ============================================================

      this._initializeDocumentTracking();


      // ============================================================
      // EXISTING PAGE USAGE TRACKING
      // ============================================================

      // Prevent duplicate log on same page load
      const pageKey = `usageLogged:${location.href}`;

      if (sessionStorage.getItem(pageKey)) {
        return;
      }


      const delay = this.properties?.ThrottleMs ?? 0;

      if (delay > 0) {
        await new Promise(resolve => setTimeout(resolve, delay));
      }


      let pageUrl = location.href.split('?')[0];


      try {

        const entries =
          performance.getEntriesByType?.('navigation');

        if (entries?.length) {

          const navEntry =
            entries[0] as PerformanceNavigationTiming;

          if (navEntry.name) {
            pageUrl =
              navEntry.name.split('?')[0];
          }

        }

      } catch {

        // fallback already set

      }


      const pageTitle =
        (document.title || '').substring(0, 255);

      const referrer =
        (document.referrer || '').substring(0, 255);

      const siteUrl =
        this.context.pageContext.site.absoluteUrl;

      const webUrl =
        this.context.pageContext.web.absoluteUrl;

      const userDisp =
        (this.context.pageContext.user.displayName || '')
          .substring(0, 255);

      const clientInfo =
        this._getClientInfo().substring(0, 255);

      const sessionId =
        this._ensureSessionId();

      const isHome =
        this._isHomePage();


      const targetSite =
        this.properties?.LoggingSiteUrl || webUrl;

      const listTitle =
        this.properties?.ListTitle || 'Site Usage Data';


      await this._addItem(
        targetSite,
        listTitle,
        {
          PageUrl: pageUrl,
          PageTitle: pageTitle,
          Referrer: referrer,
          SiteUrl: siteUrl,
          WebUrl: webUrl,
          UserDisplayName: userDisp,
          SessionId: sessionId,
          IsHomePage: isHome,
          ClientInfo: clientInfo,
          TimeStamp: new Date().toISOString(),
          RawPath: location.pathname,
          LocationHash: location.hash
        }
      );


      sessionStorage.setItem(
        pageKey,
        '1'
      );


      Log.info(
        LOG_SOURCE,
        'Usage item created.'
      );


    } catch (e) {

      Log.warn(
        LOG_SOURCE,
        e as any
      );

    }
  }



  // ==============================================================
  // DOCUMENT TRACKING
  // ==============================================================

  private _initializeDocumentTracking(): void {

    document.addEventListener(
      'click',
      (event: MouseEvent) => {

        try {

          const target =
            event.target as HTMLElement;

          if (!target) {
            return;
          }


          // Find closest link
          const anchor =
            target.closest('a') as HTMLAnchorElement;

          if (!anchor || !anchor.href) {
            return;
          }


          const documentUrl =
            anchor.href;


          // Only continue if this is a supported document
          if (!this._isDocument(documentUrl)) {
            return;
          }


          this._logDocumentInteraction(
            documentUrl
          ).catch((error) => {

            Log.warn(
              LOG_SOURCE,
              `Document logging failed: ${error}`
            );

          });


        } catch (error) {

          Log.warn(
            LOG_SOURCE,
            `Document click detection failed: ${error}`
          );

        }

      },
      true
    );

  }



  // ==============================================================
  // CHECK IF CLICKED URL IS A DOCUMENT
  // ==============================================================

  private _isDocument(url: string): boolean {

    try {

      const cleanUrl =
        url
          .split('?')[0]
          .toLowerCase();


      const extensions: string[] = [

        '.pdf',

        '.doc',
        '.docx',

        '.xls',
        '.xlsx',

        '.ppt',
        '.pptx',

        '.csv',

        '.txt',

        '.zip'

      ];


      return extensions.some(
        extension =>
          cleanUrl.endsWith(extension)
      );


    } catch {

      return false;

    }

  }



  // ==============================================================
  // LOG DOCUMENT INTERACTION
  // ==============================================================

  private async _logDocumentInteraction(
    documentUrl: string
  ): Promise<void> {

    try {

      // Remove query string
      const cleanUrl =
        documentUrl.split('?')[0];


      // Decode URL safely
      let decodedUrl =
        cleanUrl;


      try {

        decodedUrl =
          decodeURIComponent(cleanUrl);

      } catch {

        decodedUrl =
          cleanUrl;

      }


      // ============================================================
      // DOCUMENT NAME
      // ============================================================

      const documentName =
        decodedUrl.substring(
          decodedUrl.lastIndexOf('/') + 1
        );


      // ============================================================
      // FILE TYPE
      // ============================================================

      const extensionIndex =
        documentName.lastIndexOf('.');


      const fileType =
        extensionIndex > -1
          ? documentName
              .substring(extensionIndex + 1)
              .toLowerCase()
          : '';


      // ============================================================
      // SHAREPOINT CONTEXT
      // ============================================================

      const siteUrl =
        this.context.pageContext.site.absoluteUrl;


      const webUrl =
        this.context.pageContext.web.absoluteUrl;


      // ============================================================
      // USER INFORMATION
      // ============================================================

      const userDisplayName =
        (
          this.context
            .pageContext
            .user
            .displayName || ''
        ).substring(0, 255);


      const userEmail =
        (
          this.context
            .pageContext
            .user
            .email || ''
        ).substring(0, 255);


      // ============================================================
      // SESSION
      // ============================================================

      const sessionId =
        this._ensureSessionId();


      // Page where user clicked the document
      const sourcePageUrl =
        location.href.split('?')[0];


      // Central logging site
      const targetSite =
        this.properties?.LoggingSiteUrl
        || webUrl;



      // ============================================================
      // DUPLICATE PROTECTION
      //
      // SharePoint can trigger multiple browser events from one click.
      // We ignore the same document for 3 seconds.
      // ============================================================

      const duplicateKey =
        `fredDocument:${decodedUrl}`;


      const previousClick =
        sessionStorage.getItem(
          duplicateKey
        );


      const currentTime =
        Date.now();


      if (previousClick) {

        const difference =
          currentTime -
          Number(previousClick);


        if (difference < 3000) {

          return;

        }

      }


      sessionStorage.setItem(
        duplicateKey,
        currentTime.toString()
      );



      // ============================================================
      // SAVE TO DOCUMENT USAGE DATA
      // ============================================================

      await this._addItem(

        targetSite,

        'Document Usage Data',

        {

          DocumentName:
            documentName.substring(
              0,
              255
            ),

          DocumentUrl:
            decodedUrl,

          // We'll populate this in the next version
          LibraryName:
            '',

          SiteUrl:
            siteUrl,

          WebUrl:
            webUrl,

          UserDisplayName:
            userDisplayName,

          UserEmail:
            userEmail,

          FileType:
            fileType,

          ActionType:
            'DocumentClicked',

          SessionId:
            sessionId,

          SourcePageUrl:
            sourcePageUrl,

          // We'll populate this in the next version
          DocumentUniqueId:
            '',

          TimeStamp:
            new Date().toISOString()

        }

      );


      Log.info(
        LOG_SOURCE,
        `Document logged: ${documentName}`
      );


    } catch (error) {

      Log.warn(
        LOG_SOURCE,
        `Document interaction error: ${error}`
      );

    }

  }



  // ==============================================================
  // EXISTING CLIENT INFORMATION
  // ==============================================================

  private _getClientInfo(): string {

    const w: any = window;

    const s = screen;

    const dpr =
      w.devicePixelRatio || 1;


    return (
      `ua:${navigator.userAgent}` +
      `|wh:${w.innerWidth}x${w.innerHeight}` +
      `|scr:${s?.width}x${s?.height}` +
      `|dpr:${dpr}`
    );

  }



  // ==============================================================
  // EXISTING SESSION ID
  // ==============================================================

  private _ensureSessionId(): string {

    let id =
      sessionStorage.getItem(
        'usageSessionId'
      );


    if (!id) {

      id =
        Math.random()
          .toString(36)
          .slice(2)
        +
        Date.now()
          .toString(36);


      sessionStorage.setItem(
        'usageSessionId',
        id
      );

    }


    return id;

  }



  // ==============================================================
  // EXISTING HOME PAGE CHECK
  // ==============================================================

  private _isHomePage(): boolean {

    try {

      const serverRel =
        this.context
          .pageContext
          .site
          .serverRelativeUrl
          ?.replace(/\/$/, '')
        || '';


      const current =
        (
          this.context.pageContext.site as any
        )
          .serverRequestPath
          ?.replace(/\/$/, '')

        ||

        location.pathname
          .replace(/\/$/, '');


      return (
        current === '' ||
        current === '/' ||
        current === serverRel
      );


    } catch {

      return false;

    }

  }



  // ==============================================================
  // EXISTING SHAREPOINT LIST WRITER
  // ==============================================================

  private async _addItem(
    targetSiteUrl: string,
    listTitle: string,
    payload: any
  ): Promise<void> {

    const endpoint =
      `${targetSiteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items`;


    const resp: SPHttpClientResponse =
      await this.context.spHttpClient.post(

        endpoint,

        SPHttpClient.configurations.v1,

        {

          headers: {

            'Accept':
              'application/json;odata=nometadata',

            'Content-type':
              'application/json;odata=nometadata'

          },

          body:
            JSON.stringify(payload)

        }

      );


    if (!resp.ok) {

      const text =
        await resp.text();


      throw new Error(
        `Usage log failed: ` +
        `${resp.status} ` +
        `${resp.statusText} - ` +
        `${text}`
      );

    }

  }

}