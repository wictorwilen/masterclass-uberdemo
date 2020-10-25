import { BotDeclaration, PreventIframe, BotCallingWebhook, MessageExtensionDeclaration } from "express-msteams-host";
import * as debug from "debug";
import { ConfirmPrompt, Dialog, DialogSet, DialogState, DialogTurnResult, DialogTurnStatus, OAuthPrompt, WaterfallDialog, WaterfallStepContext } from "botbuilder-dialogs";
import { StatePropertyAccessor, CardFactory, TurnContext, MemoryStorage, ConversationState, ActivityTypes, TeamsActivityHandler, SigninStateVerificationQuery, BotFrameworkAdapter, UserState } from "botbuilder";
import HelpDialog from "./dialogs/HelpDialog";
import Masterclass2020MessageExtension from "../masterclass2020MessageExtension/Masterclass2020MessageExtension";
import Masterclass2020ActionsMessageExtension from "../masterclass2020ActionsMessageExtension/Masterclass2020ActionsMessageExtension";
import WelcomeCard from "./dialogs/WelcomeDialog";
import express = require("express");
import { ConfidentialClientApplication, Configuration, LogLevel } from "@azure/msal-node";
import { MainDialog } from "./dialogs/MainDialog";
import { Client } from "@microsoft/microsoft-graph-client";
// Initialize debug logging module
const log = debug("msteams");


/**
 * Implementation for masterclass2020
 */
@BotDeclaration(
    "/api/messages",
    new MemoryStorage(),
    process.env.MICROSOFT_APP_ID,
    process.env.MICROSOFT_APP_PASSWORD)
@PreventIframe("/masterclass2020Bot/faq.html")
export class Masterclass2020 extends TeamsActivityHandler {
    private readonly conversationState: ConversationState;
    /** Local property for Masterclass2020ActionsMessageExtension */
    @MessageExtensionDeclaration("masterclass2020ActionsMessageExtension")
    private _masterclass2020ActionsMessageExtension: Masterclass2020ActionsMessageExtension;
    /** Local property for Masterclass2020MessageExtension */
    @MessageExtensionDeclaration("masterclass2020MessageExtension")
    private _masterclass2020MessageExtension: Masterclass2020MessageExtension;
    private readonly dialogs: DialogSet;
    private dialogState: StatePropertyAccessor<DialogState>;
    private userState: UserState;
    private dialog: Dialog;

    /**
     * The constructor
     * @param conversationState
     */
    public constructor(conversationState: ConversationState) {
        super();
        // Message extension Masterclass2020ActionsMessageExtension
        this._masterclass2020ActionsMessageExtension = new Masterclass2020ActionsMessageExtension();

        // Message extension Masterclass2020MessageExtension
        this._masterclass2020MessageExtension = new Masterclass2020MessageExtension();

        this.userState = new UserState(new MemoryStorage());

        this.conversationState = conversationState;
        this.dialogState = conversationState.createProperty("dialogState");
        this.dialogs = new DialogSet(this.dialogState);
        this.dialogs.add(new HelpDialog("help"));


        this.dialogs.add(new MainDialog("MainDialog"));

        this.dialog = new MainDialog("MainDialog");

        // Set up the Activity processing

        this.onMessage(async (context: TurnContext, next: () => Promise<void>): Promise<void> => {
            await (this.dialog as MainDialog).run(context, this.dialogState);

            const dialogContext = await this.dialogs.createContext(context);
            const results = await dialogContext.continueDialog();
            if (results.status === DialogTurnStatus.empty) {
                switch (context.activity.type) {
                    case ActivityTypes.Message:
                        let text = TurnContext.removeRecipientMention(context.activity);
                        text = text.toLowerCase();

                        if (text.startsWith("hello")) {
                            const graphClient = Client.initWithMiddleware({
                                authProvider: {
                                    getAccessToken: (opts) => {
                                        const scopes = opts && opts.scopes ? opts.scopes : ["https://graph.microsoft.com/.default"];
                                        return new Promise<string>(async (resolve, reject) => {
                                            const botAdapter = context.adapter as BotFrameworkAdapter;
                                            const token = await botAdapter.getUserToken(context, "AAD");
                                            resolve(token.token);
                                        });
                                    }
                                },
                                debugLogging: true
                            });
                            const result = await graphClient.api("me").get();
                            await context.sendActivity(`Hello ${result.displayName}`);

                        } else if (text.startsWith("help")) {
                            const c = await this.dialogs.createContext(context);
                            await c.beginDialog("help");
                        } else if (text.startsWith("signout")) {
                            const botAdapter = context.adapter as BotFrameworkAdapter;
                            await botAdapter.signOutUser(context, "AAD");
                            await context.sendActivity("You are now signed out.");

                        } else {
                            await context.sendActivity(`I\'m terribly sorry, but my developer hasn\'t trained me to do anything yet...`);
                        }
                        break;
                    default:
                        break;
                }
            }

            await next();
        });

        this.onConversationUpdate(async (context: TurnContext): Promise<void> => {
            if (context.activity.membersAdded && context.activity.membersAdded.length !== 0) {
                for (const idx in context.activity.membersAdded) {
                    if (context.activity.membersAdded[idx].id === context.activity.recipient.id) {
                        const welcomeCard = CardFactory.adaptiveCard(WelcomeCard);
                        await context.sendActivity({ attachments: [welcomeCard] });
                    }
                }
            }
        });

        this.onMessageReaction(async (context: TurnContext): Promise<void> => {
            const added = context.activity.reactionsAdded;
            if (added && added[0]) {
                await context.sendActivity({
                    textFormat: "xml",
                    text: `That was an interesting reaction (<b>${added[0].type}</b>)`
                });
            }
        });

        this.onDialog(async (context, next) => {
            await this.conversationState.saveChanges(context, false);
            await this.userState.saveChanges(context, false);
            await next();
        });

        this.handleTeamsSigninVerifyState = async (context: TurnContext, query: SigninStateVerificationQuery): Promise<void> => {
            log("handleTeamsSigninVerifyState");
            if (context.activity.value.state === query.state) {
                log("Teams sign in verification is ok");
                await (this.dialog as MainDialog).run(context, this.dialogState);
            } else {
                log("Verification failed");
            }
        };

        this.onTokenResponseEvent(async (context, next) => {
            log("Running dialog with Token Response Event Activity.");
            await (this.dialog as MainDialog).run(context, this.dialogState);
            await next();
        });
    }


    /**
     * Webhook for incoming calls
     */
    @BotCallingWebhook("/api/calling")
    public async onIncomingCall(req: express.Request, res: express.Response) {
        log("Incoming call");
        // TODO: Implement authorization header validation

        // TODO: Add your management of calls (answer, reject etc.)

        // default, send an access denied
        res.sendStatus(401);
    }
}
