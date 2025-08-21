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
import { PlannerService } from './services/plannerService';
import { PlannerCards } from './cards/plannerCards';

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
    defaultPrompt: 'planner'
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
                text: 'Please sign in to use the Planner bot.',
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
        "Hi! I'm your AI-powered Planner assistant! I can help you with natural language requests like:\n• 'Show me all my plans'\n• 'Create a new plan for marketing'\n• 'What tasks are in my project plan?'\n• 'Add a task to review the presentation'\n• 'Show me all my assigned tasks'\n\nJust tell me what you'd like to do with your plans and tasks!"
    );
});

// Register AI action handlers
app.ai.action('getPlans', async (context: TurnContext, state: ApplicationTurnState) => {
    const token = await app.getTokenOrStartSignIn(context, state, 'graph');
    if (!token) {
        await context.sendActivity('You have to be signed in to view your plans. Starting sign in flow...');
        return AI.StopCommandName;
    }

    try {
        const plannerService = new PlannerService(token);
        const plans = await plannerService.getPlans();
        const card = PlannerCards.createPlansListCard(plans);
        
        await context.sendActivity({
            attachments: [CardFactory.adaptiveCard(card)]
        });
        
        return `Found ${plans.length} plans. Displayed them in a card above.`;
    } catch (error) {
        console.error('Error getting plans:', error);
        return `Error retrieving plans: ${error instanceof Error ? error.message : 'Unknown error'}`;
    }
});

app.ai.action('createPlan', async (context: TurnContext, state: ApplicationTurnState, parameters: any) => {
    const token = await app.getTokenOrStartSignIn(context, state, 'graph');
    if (!token) {
        await context.sendActivity('You have to be signed in to create plans. Starting sign in flow...');
        return AI.StopCommandName;
    }

    try {
        const { title } = parameters;
        const plannerService = new PlannerService(token);
        
        const groups = await plannerService.getUserGroups();
        if (groups.length === 0) {
            return 'You need to be a member of a Microsoft 365 group to create plans.';
        }

        const firstGroup = groups[0];
        const plan = await plannerService.createPlan(title, firstGroup.id);
        
        return `✅ Plan "${plan.title}" created successfully! Plan ID: ${plan.id}`;
    } catch (error) {
        console.error('Error creating plan:', error);
        return `Error creating plan: ${error instanceof Error ? error.message : 'Unknown error'}`;
    }
});

app.ai.action('getPlanTasks', async (context: TurnContext, state: ApplicationTurnState, parameters: any) => {
    const token = await app.getTokenOrStartSignIn(context, state, 'graph');
    if (!token) {
        await context.sendActivity('You have to be signed in to view tasks. Starting sign in flow...');
        return AI.StopCommandName;
    }

    try {
        const { planId } = parameters;
        const plannerService = new PlannerService(token);
        
        // Try to get all data, but handle permissions issues gracefully
        let tasks: any[] = [];
        let buckets: any[] = [];
        let planTitle = 'Plan';
        
        try {
            const [tasksResult, plans] = await Promise.all([
                plannerService.getPlanTasks(planId),
                plannerService.getPlans()
            ]);
            
            tasks = tasksResult;
            const plan = plans.find(p => p.id === planId);
            planTitle = plan ? plan.title : 'Plan';
            
            // Try to get buckets separately to handle permissions
            try {
                buckets = await plannerService.getPlanBuckets(planId);
            } catch (bucketsError) {
                console.warn('Could not retrieve buckets, will display tasks without bucket organization');
                buckets = []; // Empty buckets array, tasks will be shown in a simple list
            }
            
        } catch (tasksError: any) {
            if (tasksError.message.includes('permissions') || tasksError.message.includes('Forbidden')) {
                return 'I cannot access the tasks in this plan due to insufficient permissions. Please make sure you have the necessary Planner permissions to view tasks in this plan.';
            }
            throw tasksError; // Re-throw if it's not a permissions error
        }

        const card = PlannerCards.createTasksBoardCard(tasks, buckets, planTitle);
        
        await context.sendActivity({
            attachments: [CardFactory.adaptiveCard(card)]
        });
        
        return `Found ${tasks.length} tasks in plan "${planTitle}". Displayed them in a card above.`;
    } catch (error) {
        console.error('Error getting plan tasks:', error);
        return `Error retrieving plan tasks: ${error instanceof Error ? error.message : 'Unknown error'}`;
    }
});

app.ai.action('addTask', async (context: TurnContext, state: ApplicationTurnState, parameters: any) => {
    const token = await app.getTokenOrStartSignIn(context, state, 'graph');
    if (!token) {
        await context.sendActivity('You have to be signed in to add tasks. Starting sign in flow...');
        return AI.StopCommandName;
    }

    try {
        const { planId, title } = parameters;
        const plannerService = new PlannerService(token);
        
        // Try to get buckets, but handle permissions gracefully
        let buckets: any[] = [];
        try {
            buckets = await plannerService.getPlanBuckets(planId);
        } catch (bucketsError) {
            console.error('Could not retrieve buckets:', bucketsError);
            // If we can't get buckets, try to create the task without specifying a bucket
            // and let Microsoft Planner assign it to the default bucket
        }

        if (buckets.length === 0) {
            // Try to create task without bucket (will go to default bucket)
            try {
                const me = await plannerService.getCurrentUser();
                const task = await plannerService.createTaskWithoutBucket(planId, title, me.id);
                return `✅ Task "${task.title}" created successfully in the default bucket!`;
            } catch (createError) {
                return 'I cannot add tasks to this plan due to insufficient permissions. Please make sure you have the necessary Planner permissions and that the plan has at least one bucket created.';
            }
        }

        // Get current user for task assignment
        const me = await plannerService.getCurrentUser();
        const task = await plannerService.createTask(planId, buckets[0].id, title, me.id);
        
        return `✅ Task "${task.title}" created successfully in the plan!`;
    } catch (error) {
        console.error('Error adding task:', error);
        return `Error adding task: ${error instanceof Error ? error.message : 'Unknown error'}`;
    }
});

app.ai.action('getMyTasks', async (context: TurnContext, state: ApplicationTurnState) => {
    const token = await app.getTokenOrStartSignIn(context, state, 'graph');
    if (!token) {
        await context.sendActivity('You have to be signed in to view your tasks. Starting sign in flow...');
        return AI.StopCommandName;
    }

    try {
        const plannerService = new PlannerService(token);
        const tasks = await plannerService.getMyTasks();
        const card = PlannerCards.createMyTasksCard(tasks);
        
        await context.sendActivity({
            attachments: [CardFactory.adaptiveCard(card)]
        });
        
        return `Found ${tasks.length} tasks assigned to you. Displayed them in a card above.`;
    } catch (error) {
        console.error('Error getting my tasks:', error);
        return `Error retrieving your tasks: ${error instanceof Error ? error.message : 'Unknown error'}`;
    }
});

// Handle signout command
app.message('/signout', async (context: TurnContext, state: ApplicationTurnState) => {
    await app.authentication.signOutUser(context, state);
    await context.sendActivity('You have signed out successfully.');
});

// Authentication event handlers
app.authentication.get('graph').onUserSignInSuccess(async (context: TurnContext) => {
    await context.sendActivity('Thanks for signing in! I can now help you with your Planner tasks and plans.');
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