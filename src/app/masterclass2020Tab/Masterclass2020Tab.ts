import { PreventIframe } from "express-msteams-host";

/**
 * Used as place holder for the decorators
 */
@PreventIframe("/masterclass2020Tab/index.html")
@PreventIframe("/masterclass2020Tab/config.html")
@PreventIframe("/masterclass2020Tab/remove.html")
export class Masterclass2020Tab {
}
