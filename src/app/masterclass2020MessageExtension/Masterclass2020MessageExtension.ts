import * as debug from "debug";
import { PreventIframe } from "express-msteams-host";
import { TurnContext, CardFactory, MessagingExtensionQuery, MessagingExtensionResult } from "botbuilder";
import { IMessagingExtensionMiddlewareProcessor } from "botbuilder-teams-messagingextensions";
import * as users from "../users.json";
// Initialize debug logging module
const log = debug("msteams");

@PreventIframe("/masterclass2020MessageExtension/config.html")
export default class Masterclass2020MessageExtension implements IMessagingExtensionMiddlewareProcessor {

    public generateCard(user) {
        const card = CardFactory.adaptiveCard(
            {
                type: "AdaptiveCard",
                body: [
                    {
                        type: "TextBlock",
                        size: "Large",
                        text: user.name.first + " " + user.name.last
                    },
                    {
                        type: "TextBlock",
                        text: user.location.country
                    },
                    {
                        type: "TextBlock",
                        text: user.email
                    },
                    {
                        type: "Image",
                        url: user.picture.large
                    }
                ],
                actions: [
                    {
                        type: "Action.Submit",
                        title: "More details",
                        data: {
                            action: "moreDetails",
                            id: user.id.value
                        }
                    }
                ],
                $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
                version: "1.0"
            });
        const preview = {
            contentType: "application/vnd.microsoft.card.thumbnail",
            content: {
                title: user.name.first + " " + user.name.last,
                text: user.location.country,
                images: [
                    {
                        url: user.picture.thumbnail
                    }
                ]
            }
        };
        return { ...card, preview };
    }




    public async onQuery(context: TurnContext, query: MessagingExtensionQuery): Promise<MessagingExtensionResult> {

        if (query.parameters && query.parameters[0] && query.parameters[0].name === "initialRun") {
            // initial run
            const result = users.results.slice(10).map(u => {
                return this.generateCard(u);
            });

            return Promise.resolve({
                type: "result",
                attachmentLayout: "list",
                attachments: result
            } as MessagingExtensionResult);
        } else {
            // the rest
            const result = users.results.filter(x => {
                return query.parameters && query.parameters[0] && (
                    x.name.first.toLowerCase().startsWith(query.parameters[0].value.toLowerCase()) ||
                    x.name.last.toLowerCase().startsWith(query.parameters[0].value.toLowerCase())
                );
            }).map(u => {
                return this.generateCard(u);
            });
            return Promise.resolve({
                type: "result",
                attachmentLayout: "list",
                attachments: result
            } as MessagingExtensionResult);
        }
    }

    public async onCardButtonClicked(context: TurnContext, value: any): Promise<void> {
        // Handle the Action.Submit action on the adaptive card
        if (value.action === "moreDetails") {
            log(`I got this ${value.id}`);
        }
        return Promise.resolve();
    }







    // this is used when canUpdateConfiguration is set to true
    public async onQuerySettingsUrl(context: TurnContext): Promise<{ title: string, value: string }> {
        return Promise.resolve({
            title: "Masterclass 2020 Configuration",
            value: `https://${process.env.HOSTNAME}/masterclass2020MessageExtension/config.html?name={loginHint}&tenant={tid}&group={groupId}&theme={theme}`
        });
    }

    public async onSettings(context: TurnContext): Promise<void> {
        // take care of the setting returned from the dialog, with the value stored in state
        const setting = context.activity.value.state;
        log(`New setting: ${setting}`);
        return Promise.resolve();
    }

}
