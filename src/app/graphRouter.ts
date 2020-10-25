import express = require("express");
import passport = require("passport");
import { BearerStrategy, VerifyCallback, IBearerStrategyOption, ITokenPayload } from "passport-azure-ad";
import qs = require("querystring");
import Axios from "axios";
import * as debug from "debug";
import { ConfidentialClientApplication, Configuration, LogLevel } from "@azure/msal-node";
import { Client } from "@microsoft/microsoft-graph-client";
const log = debug("graphRouter");

export const graphRouter = (options: any): express.Router => {
    const router = express.Router();

    const config: Configuration = {
        auth: {
            clientId: process.env.MASTERCLASS2020_APP_ID as string,
            clientSecret: process.env.MASTERCLASS2020_APP_SECRET
        },
        cache: {
            // cachePlugin: TeamsGraph.cache
        },
        system: {
            loggerOptions: {
                loggerCallback(logLevel: LogLevel, message: string, containsPii: boolean) {
                    log(message);
                },
                piiLoggingEnabled: false,
                logLevel: LogLevel.Verbose,
            }
        }
    };
    const cca = new ConfidentialClientApplication(config);

    // Set up the Bearer Strategy
    const bearerStrategy = new BearerStrategy({
        identityMetadata: "https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration",
        clientID: process.env.MASTERCLASS2020_APP_ID as string,
        audience: process.env.MASTERCLASS2020_APP_ID as string,
        loggingLevel: "warn",
        validateIssuer: false,
        passReqToCallback: false
    } as IBearerStrategyOption,
        (token: ITokenPayload, done: VerifyCallback) => {
            done(null, { tid: token.tid, name: token.name, upn: token.upn }, token);
        }
    );
    const pass = new passport.Passport();
    router.use(pass.initialize());
    pass.use(bearerStrategy);


    // Define the rout for the photo
    router.get(
        "/photo",
        pass.authenticate("oauth-bearer", { session: false }),
        async (req: express.Request, res: express.Response, next: express.NextFunction) => {
            const user: any = req.user;
            try {
                const assertion = req.header("Authorization")!.replace("Bearer ", "") as string;

                const graphClient = Client.initWithMiddleware({
                    authProvider: {
                        getAccessToken: (opts) => {
                            const scopes = opts && opts.scopes ? opts.scopes : ["https://graph.microsoft.com/.default"];
                            return new Promise<string>((resolve, reject) => {
                                cca.acquireTokenOnBehalfOf({
                                    scopes,
                                    authority: "https://login.microsoftonline.com/" + user.tid,
                                    oboAssertion: assertion,
                                    skipCache: true
                                }).then((response) => {
                                    resolve(response.accessToken);
                                }).catch((error) => {
                                    reject(error);
                                });
                            });

                        }
                    },
                    debugLogging: true
                });

                graphClient.api("me/photo/$value").get(async (err, response: Blob, rawResponse) => {
                    if (!err) {
                        res.type(response.type);
                        res.status(200);
                        res.end(Buffer.from(await response.arrayBuffer()), "binary");
                    } else {
                        res.status(500).send(err);
                    }
                });

            } catch (err) {
                if (err.status) {
                    res.status(err.status).send(err.message);
                } else {
                    res.status(500).send(err);
                }
            }

        });
    return router;
};
