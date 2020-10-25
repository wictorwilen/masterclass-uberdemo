import { TurnContext, StatePropertyAccessor } from "botbuilder";
import { ComponentDialog, DialogSet, DialogState, DialogTurnResult, DialogTurnStatus, OAuthPrompt, WaterfallDialog, WaterfallStepContext } from "botbuilder-dialogs";
import * as debug from "debug";
const log = debug("msteams");

const loginPrompt = new OAuthPrompt("LoginDialog", {
    connectionName: "AAD",
    text: "I need you to sign in first",
    title: "Login",
    timeout: 30000 // User has 5 minutes to login.
});

export class MainDialog extends ComponentDialog {
    constructor(id: string) {
        super(id);

        this.addDialog(loginPrompt)
            .addDialog(new WaterfallDialog("MainDialog", [
                this.promptStep.bind(this),
                this.initialStep.bind(this)
            ]));
        this.initialDialogId = "MainDialog";
    }

    public async run(context: TurnContext, accessor: StatePropertyAccessor<DialogState>) {
        const dialogSet = new DialogSet(accessor);
        dialogSet.add(this);

        const dialogContext = await dialogSet.createContext(context);
        const results = await dialogContext.continueDialog();
        if (results.status === DialogTurnStatus.empty) {
            await dialogContext.beginDialog(this.id);
        }
    }

    private async promptStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
        return await stepContext.beginDialog("LoginDialog");
    }

    private async initialStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
        const tokenResponse = stepContext.result;
        if (tokenResponse) {
            log(tokenResponse);
            log("User logged in.");
            return await stepContext.endDialog();
        }
        log("Login failed!");
        await stepContext.context.sendActivity("Ouch! I could not sign you in, mayhaps I mixed up some 1's and 0's?");
        return await stepContext.endDialog();
    }
}

