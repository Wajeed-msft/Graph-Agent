import { config } from 'dotenv';
import * as path from 'path';
import * as restify from 'restify';
import {
    ActivityTypes,
    ConfigurationServiceClientCredentialFactory,
    MemoryStorage,
    TurnContext,
    CardFactory
} from 'botbuilder';
import {
    AI,
    ActionPlanner,
    ApplicationBuilder,
    DefaultConversationState,
    OpenAIModel,
    PromptManager,
    TeamsAdapter,
    TurnState
} from '@microsoft/teams-ai';
import { {WORKLOAD_NAME}Service } from './services/graphService.template';
import { {WORKLOAD_NAME}Cards } from './cards/graphCards.template';

// Read environment variables
config();

if (!process.env.OPENAI_API_KEY) {
    throw new Error('Missing environment variables - please check that OPENAI_API_KEY is set.');
}

// Create adapter
const adapter = new TeamsAdapter(
    {},
    new ConfigurationServiceClientCredentialFactory({
        MicrosoftAppId: process.env.CLIENT_ID || process.env.BOT_ID,
        MicrosoftAppPassword: process.env.CLIENT_SECRET || process.env.BOT_PASSWORD,
        MicrosoftAppType: 'MultiTenant'
    })
);

// Catch-all for errors
const onTurnErrorHandler = async (context: TurnContext, error: Error) => {
    console.error(`\n [onTurnError] unhandled error: ${error.toString()}`);

    await context.sendTraceActivity(
        'OnTurnError Trace',
        `${error.toString()}`,
        'https://www.botframework.com/schemas/error',
        'TurnError'
    );

    await context.sendActivity('The bot encountered an error or bug.');
    await context.sendActivity('To continue to run this bot, please fix the bot source code.');
};

// Set the onTurnError for the singleton CloudAdapter
adapter.onTurnError = onTurnErrorHandler;

// Create HTTP server
const server = restify.createServer();
server.use(restify.plugins.bodyParser());

const port = process.env.PORT || 3978;
server.listen(port, () => {
    console.log(`\n${server.name} listening to ${server.url}`);
    console.log('\nTo test your bot in Teams, sideload the app manifest.json within Teams Apps.');
});

// Create AI components
const model = new OpenAIModel({
    apiKey: process.env.OPENAI_API_KEY!,
    defaultModel: 'gpt-4o',
    logRequests: true
});

const prompts = new PromptManager({
    promptsFolder: path.join(__dirname, 'prompts')
});

const planner = new ActionPlanner({
    model,
    prompts,
    defaultPrompt: 'generic'
});

// Define application turn state
export type ApplicationTurnState = TurnState<DefaultConversationState>;

// Define storage and application
const storage = new MemoryStorage();
const app = new ApplicationBuilder<ApplicationTurnState>()
    .withStorage(storage)
    .withAuthentication(adapter, {
        settings: {
            graph: {
                connectionName: process.env.OAUTH_CONNECTION_NAME ?? 'graph',
                title: 'Sign in',
                text: 'Please sign in to use the {WORKLOAD_NAME} bot.',
                endOnInvalidMessage: true
            }
        }
    })
    .withAIOptions({
        planner
    })
    .build();

// Register activity handlers
app.activity(ActivityTypes.InstallationUpdate, async (context: TurnContext) => {
    await context.sendActivity(
        "Hi! I'm your AI-powered {WORKLOAD_NAME} assistant! I can help you with natural language requests like:\n" +
        "• 'Show me all my {WORKLOAD_NAME_LOWER}s'\n" +
        "• 'Create a new {WORKLOAD_NAME_LOWER} for {WORKLOAD_EXAMPLE_NAME}'\n" +
        "• 'Get details for my {WORKLOAD_NAME_LOWER} item'\n" +
        "• 'Update my {WORKLOAD_NAME_LOWER} with new information'\n" +
        "• 'Show me my {WORKLOAD_NAME_LOWER} items'\n\n" +
        // TODO: Add workload-specific examples
        // Example for Calendar: "• 'What's on my calendar today?'\n• 'Schedule a meeting for tomorrow'\n"
        // Example for Teams: "• 'What teams am I in?'\n• 'Send a message to the team'\n"
        // Example for SharePoint: "• 'Show me my recent files'\n• 'Upload a document'\n"
        "Just tell me what you'd like to do with your {WORKLOAD_NAME_LOWER}s!"
    );
});

// Register AI action handlers

// Get all workload items
app.ai.action('get{WORKLOAD_NAME}s', async (context: TurnContext, state: ApplicationTurnState, parameters: any) => {
    const token = await app.getTokenOrStartSignIn(context, state, 'graph');
    if (!token) {
        await context.sendActivity('You have to be signed in to view your {WORKLOAD_NAME_LOWER}s. Starting sign in flow...');
        return AI.StopCommandName;
    }

    try {
        const service = new {WORKLOAD_NAME}Service(token);
        const items = await service.get{WORKLOAD_NAME}s(parameters);
        const card = {WORKLOAD_NAME}Cards.create{WORKLOAD_NAME}sListCard(items);

        await context.sendActivity({
            attachments: [CardFactory.adaptiveCard(card)]
        });

        return `Found ${items.length} {WORKLOAD_NAME_LOWER}s. Displayed them in a card above.`;
    } catch (error) {
        console.error('Error getting {WORKLOAD_NAME_LOWER}s:', error);
        return `Error retrieving {WORKLOAD_NAME_LOWER}s: ${error instanceof Error ? error.message : 'Unknown error'}`;
    }
});

// Create new workload item
app.ai.action('create{WORKLOAD_NAME}', async (context: TurnContext, state: ApplicationTurnState, parameters: any) => {
    const token = await app.getTokenOrStartSignIn(context, state, 'graph');
    if (!token) {
        await context.sendActivity('You have to be signed in to create {WORKLOAD_NAME_LOWER}s. Starting sign in flow...');
        return AI.StopCommandName;
    }

    try {
        const service = new {WORKLOAD_NAME}Service(token);
        const item = await service.create{WORKLOAD_NAME}(parameters);
        const card = {WORKLOAD_NAME}Cards.create{WORKLOAD_NAME}CreatedCard(item);

        await context.sendActivity({
            attachments: [CardFactory.adaptiveCard(card)]
        });

        return `✅ {WORKLOAD_NAME} "${item.{WORKLOAD_PRIMARY_FIELD}}" created successfully!`;
    } catch (error) {
        console.error('Error creating {WORKLOAD_NAME_LOWER}:', error);
        return `Error creating {WORKLOAD_NAME_LOWER}: ${error instanceof Error ? error.message : 'Unknown error'}`;
    }
});

// Get workload item details
app.ai.action('get{WORKLOAD_NAME}Details', async (context: TurnContext, state: ApplicationTurnState, parameters: any) => {
    const token = await app.getTokenOrStartSignIn(context, state, 'graph');
    if (!token) {
        await context.sendActivity('You have to be signed in to view {WORKLOAD_NAME_LOWER} details. Starting sign in flow...');
        return AI.StopCommandName;
    }

    try {
        const { id } = parameters;
        const service = new {WORKLOAD_NAME}Service(token);
        const item = await service.get{WORKLOAD_NAME}ById(id);
        const card = {WORKLOAD_NAME}Cards.create{WORKLOAD_NAME}DetailCard(item);

        await context.sendActivity({
            attachments: [CardFactory.adaptiveCard(card)]
        });

        return `Found details for {WORKLOAD_NAME_LOWER} "${item.{WORKLOAD_PRIMARY_FIELD}}". Displayed in a card above.`;
    } catch (error) {
        console.error('Error getting {WORKLOAD_NAME_LOWER} details:', error);
        return `Error retrieving {WORKLOAD_NAME_LOWER} details: ${error instanceof Error ? error.message : 'Unknown error'}`;
    }
});

// Update workload item
app.ai.action('update{WORKLOAD_NAME}', async (context: TurnContext, state: ApplicationTurnState, parameters: any) => {
    const token = await app.getTokenOrStartSignIn(context, state, 'graph');
    if (!token) {
        await context.sendActivity('You have to be signed in to update {WORKLOAD_NAME_LOWER}s. Starting sign in flow...');
        return AI.StopCommandName;
    }

    try {
        const { id, ...updates } = parameters;
        const service = new {WORKLOAD_NAME}Service(token);
        const item = await service.update{WORKLOAD_NAME}(id, updates);

        return `✅ {WORKLOAD_NAME} "${item.{WORKLOAD_PRIMARY_FIELD}}" updated successfully!`;
    } catch (error) {
        console.error('Error updating {WORKLOAD_NAME_LOWER}:', error);
        return `Error updating {WORKLOAD_NAME_LOWER}: ${error instanceof Error ? error.message : 'Unknown error'}`;
    }
});

// Delete workload item
app.ai.action('delete{WORKLOAD_NAME}', async (context: TurnContext, state: ApplicationTurnState, parameters: any) => {
    const token = await app.getTokenOrStartSignIn(context, state, 'graph');
    if (!token) {
        await context.sendActivity('You have to be signed in to delete {WORKLOAD_NAME_LOWER}s. Starting sign in flow...');
        return AI.StopCommandName;
    }

    try {
        const { id } = parameters;
        const service = new {WORKLOAD_NAME}Service(token);
        await service.delete{WORKLOAD_NAME}(id);

        return `✅ {WORKLOAD_NAME} deleted successfully!`;
    } catch (error) {
        console.error('Error deleting {WORKLOAD_NAME_LOWER}:', error);
        return `Error deleting {WORKLOAD_NAME_LOWER}: ${error instanceof Error ? error.message : 'Unknown error'}`;
    }
});

// Get user-specific workload items
app.ai.action('getMy{WORKLOAD_NAME}s', async (context: TurnContext, state: ApplicationTurnState) => {
    const token = await app.getTokenOrStartSignIn(context, state, 'graph');
    if (!token) {
        await context.sendActivity('You have to be signed in to view your {WORKLOAD_NAME_LOWER}s. Starting sign in flow...');
        return AI.StopCommandName;
    }

    try {
        const service = new {WORKLOAD_NAME}Service(token);
        const items = await service.getMy{WORKLOAD_NAME}s();
        const card = {WORKLOAD_NAME}Cards.createMy{WORKLOAD_NAME}sCard(items);

        await context.sendActivity({
            attachments: [CardFactory.adaptiveCard(card)]
        });

        return `Found ${items.length} {WORKLOAD_NAME_LOWER}s assigned to you. Displayed them in a card above.`;
    } catch (error) {
        console.error('Error getting my {WORKLOAD_NAME_LOWER}s:', error);
        return `Error retrieving your {WORKLOAD_NAME_LOWER}s: ${error instanceof Error ? error.message : 'Unknown error'}`;
    }
});

// TODO: Add workload-specific action handlers
// Example for Calendar:
// app.ai.action('getUpcomingEvents', async (context, state, parameters) => { ... });
// app.ai.action('createMeeting', async (context, state, parameters) => { ... });

// Example for Teams:
// app.ai.action('getTeamChannels', async (context, state, parameters) => { ... });
// app.ai.action('sendMessageToChannel', async (context, state, parameters) => { ... });

// Example for SharePoint:
// app.ai.action('getListItems', async (context, state, parameters) => { ... });
// app.ai.action('uploadFile', async (context, state, parameters) => { ... });

// Handle signout command
app.message('/signout', async (context: TurnContext, state: ApplicationTurnState) => {
    await app.authentication.signOutUser(context, state);
    await context.sendActivity('You have signed out successfully.');
});

// Authentication event handlers
app.authentication.get('graph').onUserSignInSuccess(async (context: TurnContext) => {
    await context.sendActivity('Thanks for signing in! I can now help you with your {WORKLOAD_NAME} items.');
});

app.authentication.get('graph').onUserSignInFailure(async (context: TurnContext, _state: ApplicationTurnState, error: any) => {
    await context.sendActivity('Failed to sign in. Please try again.');
    await context.sendActivity(`Error: ${error.message}`);
});

// Listen for incoming server requests
server.post('/api/messages', async (req: restify.Request, res: restify.Response) => {
    try {
        await adapter.process(req, res as any, async (context) => {
            await app.run(context);
        });
    } catch (error) {
        console.error('Error processing message:', error);
        res.status(500);
        res.send('Internal server error');
    }
});

// Add some basic error handling for startup
process.on('unhandledRejection', (reason, promise) => {
    console.error('Unhandled Rejection at:', promise, 'reason:', reason);
});

process.on('uncaughtException', (error) => {
    console.error('Uncaught Exception:', error);
    process.exit(1);
});