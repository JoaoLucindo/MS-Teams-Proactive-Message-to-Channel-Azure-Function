import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { Activity, BotFrameworkAdapter, CardFactory, ConversationParameters, MessageFactory } from "botbuilder";
import { MicrosoftAppCredentials } from "botframework-connector";

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {

    // Initialize Connector Client
    const adapter = new BotFrameworkAdapter({
        appId: process.env["MICROSOFT_APP_ID"],
        appPassword: process.env["MICROSOFT_APP_PASSWORD"]
    });

    MicrosoftAppCredentials.trustServiceUrl(process.env["SERVICE_URL"]);

    const connectorClient = adapter.createConnectorClient(process.env["SERVICE_URL"]);

    // Init req.body parameters
    const card = req.body.card;
    const tenantid = process.env["TENANT_ID"];
    const proactiveCard = CardFactory.adaptiveCard(card);
    const message = MessageFactory.attachment(proactiveCard) as Activity;

    // Parse if req body contains user or channel
    switch(req.body.destination) {

        case "user":            
            // Send card to user
            const userid = req.body.userid;

            // User Scope
            const userConversationParameters = {
                isGroup: false,
                channelData: {
                    tenant: {
                        id: tenantid
                    }
                },
                bot: {
                    id: process.env["BOT_ID"],
                    name: process.env["BOT_NAME"]
                },
                members: [
                    {
                        id: userid
                    }
                ]
            } as ConversationParameters;

            // Leverage Connector Client to send message
            const response = await connectorClient.conversations.createConversation(userConversationParameters);
            await connectorClient.conversations.sendToConversation(response.id, message);

            // Send response
            context.res = {
                // status: 200, /* Defaults to 200 */
                body: "Mensagem enviada com sucesso para o usuário"
            };


            break;
        case "channel":
            // Send card to channel
            const channelid = req.body.channelid;

            // Channel Scope
            const channelConversationParameters = {
                isGroup: true,
                channelData: {
                    channel: {
                        id: channelid
                    }
                },
                activity: message
            } as ConversationParameters;

            // Leverage Connector Client to send message
            await connectorClient.conversations.createConversation(channelConversationParameters);

            // Send response
            context.res = {
                // status: 200, /* Defaults to 200 */
                body: "Mensagem enviada com sucesso para o canal"
            };


            break;
        default:
            break;
    }
    

};

export default httpTrigger;