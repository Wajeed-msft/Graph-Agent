# M365 Graph Agent Template - Claude Code Implementation Guide

## Overview
This is a template for creating Microsoft Graph API agents that can be customized for any M365 workload (Calendar, Teams, SharePoint, OneDrive, etc.). The template provides a generic structure that Claude Code can automatically customize based on user prompts.

## Template Structure

### Core Template Files
- `src/types/graphTypes.template.ts` - Generic type definitions with placeholders
- `src/services/baseGraphService.ts` - Common Graph API functionality
- `src/services/graphService.template.ts` - Template service class
- `src/cards/graphCards.template.ts` - Template card components
- `src/index.template.ts` - Main application template
- `src/prompts/generic/` - Template prompt files

### Working Reference Implementation
The existing Planner implementation serves as a reference:
- `src/services/plannerService.ts` - Working example
- `src/cards/plannerCards.ts` - Working card examples
- `src/prompts/planner/` - Working prompt configuration

## Implementation Guide for Claude Code

### Step 1: Identify Workload from User Prompt
User prompts will be like:
- "Build me a calendar tracking agent"
- "Create a Teams management bot"
- "I need a SharePoint file manager"

Extract the workload name and map to Graph API resources.

### Step 2: Template Placeholder Replacement

Replace these placeholders throughout the template files:

| Placeholder | Example (Calendar) | Example (Teams) | Example (SharePoint) |
|-------------|-------------------|-----------------|---------------------|
| `{WORKLOAD_NAME}` | Calendar | Team | Site |
| `{WORKLOAD_NAME_LOWER}` | calendar | team | site |
| `{WORKLOAD_NAME_UPPER}` | CALENDAR | TEAM | SITE |
| `{WORKLOAD_ICON}` | =Å | =e | =Á |
| `{WORKLOAD_PRIMARY_FIELD}` | subject | displayName | name |
| `{WORKLOAD_PRIMARY_FIELD_DESCRIPTION}` | title/subject of the event | name of the team | name of the site |
| `{WORKLOAD_BASE_ENDPOINT}` | /me/events | /teams | /sites |
| `{WORKLOAD_CREATE_ENDPOINT}` | /me/events | /teams | /sites |
| `{WORKLOAD_USER_ENDPOINT}` | events | joinedTeams | followedSites |
| `{WORKLOAD_EXAMPLE_NAME}` | Marketing Meeting | Project Alpha | Document Library |

### Step 3: Implement Workload-Specific Logic

#### For Calendar Workload:
```typescript
// In graphTypes.template.ts
interface CalendarItem {
  id: string;
  subject: string;
  start: { dateTime: string; timeZone: string };
  end: { dateTime: string; timeZone: string };
  body?: { content: string };
  attendees?: Array<{ emailAddress: { address: string; name: string } }>;
  location?: { displayName: string };
}

// Endpoints
BASE: '/me/events'
CREATE: '/me/events'

// Permissions needed
CALENDAR_PERMISSIONS = {
  READ: 'Calendars.Read',
  WRITE: 'Calendars.ReadWrite'
}
```

#### For Teams Workload:
```typescript
// In graphTypes.template.ts
interface TeamItem {
  id: string;
  displayName: string;
  description?: string;
  visibility: 'public' | 'private';
  memberSettings?: any;
}

// Endpoints
BASE: '/teams'
USER_ENDPOINT: 'joinedTeams'

// Permissions needed
TEAM_PERMISSIONS = {
  READ: 'Team.ReadBasic.All',
  WRITE: 'Team.Create'
}
```

### Step 4: Update Infrastructure Files

#### Azure Bot Registration (`infra/botRegistration/azurebot.bicep`)
Update the `properties.scopes` string with workload-specific permissions:

**Calendar**: `Calendars.Read,Calendars.ReadWrite,User.Read`
**Teams**: `Team.ReadBasic.All,Team.Create,User.Read`
**SharePoint**: `Sites.Read.All,Sites.ReadWrite.All,User.Read`

#### AAD Manifest (`aad.manifest.json`)
Update `requiredResourceAccess` with corresponding permission IDs:

```json
"requiredResourceAccess": [
  {
    "resourceAppId": "00000003-0000-0000-c000-000000000000",
    "resourceAccess": [
      {
        "id": "465a38f9-76ea-45b9-9f34-9e8b0d4b0b42", // Calendars.Read
        "type": "Scope"
      }
    ]
  }
]
```

### Step 5: Common Graph API Patterns

#### CRUD Operations Pattern
```typescript
// GET all items
async getItems(): Promise<Item[]> {
  return await this.getPaginatedResults<Item>('/endpoint');
}

// GET by ID
async getById(id: string): Promise<Item> {
  return await this.graphClient.api(`/endpoint/${id}`).get();
}

// CREATE item
async create(data: CreateRequest): Promise<Item> {
  return await this.createResource<Item>('/endpoint', data);
}

// UPDATE item
async update(id: string, data: UpdateRequest): Promise<Item> {
  return await this.updateResource<Item>(`/endpoint/${id}`, data);
}

// DELETE item
async delete(id: string): Promise<void> {
  await this.deleteResource(`/endpoint/${id}`);
}
```

#### Error Handling Pattern
```typescript
try {
  // Graph API call
} catch (error: any) {
  throw this.handleGraphError(error, 'operation description');
}
```

### Step 6: Workload-Specific Actions

Add custom actions based on the workload:

#### Calendar-Specific Actions:
- `getUpcomingEvents` - Get events for next N days
- `createMeeting` - Create meeting with attendees
- `respondToInvitation` - Accept/decline meeting invites

#### Teams-Specific Actions:
- `getTeamChannels` - Get channels for a team
- `sendMessageToChannel` - Send message to team channel
- `createChannel` - Create new channel in team

#### SharePoint-Specific Actions:
- `getListItems` - Get items from SharePoint list
- `uploadFile` - Upload file to document library
- `shareFile` - Share file with users

### Step 7: File Renaming and Organization

After implementing the workload-specific logic:

1. Copy template files to workload-specific names:
   - `graphService.template.ts` ’ `calendarService.ts`
   - `graphCards.template.ts` ’ `calendarCards.ts`
   - `index.template.ts` ’ `index.ts` (replace existing)

2. Update import statements to reference the new files

3. Update the prompts folder:
   - Copy `src/prompts/generic/` ’ `src/prompts/calendar/`
   - Update prompt manager in index.ts: `defaultPrompt: 'calendar'`

## Permission Reference Guide

### Common Permissions by Workload

| Workload | Read Permission | Write Permission | Additional |
|----------|----------------|------------------|------------|
| Calendar | Calendars.Read | Calendars.ReadWrite | Calendars.ReadWrite.Shared |
| Teams | Team.ReadBasic.All | Team.Create | Channel.Create, TeamMember.ReadWrite.All |
| SharePoint | Sites.Read.All | Sites.ReadWrite.All | Files.ReadWrite.All |
| OneDrive | Files.Read | Files.ReadWrite | Files.ReadWrite.All |
| Mail | Mail.Read | Mail.Send | Mail.ReadWrite |
| Users | User.Read.All | User.ReadWrite.All | Directory.Read.All |

### Permission ID Reference
Check Microsoft Graph permissions documentation for the exact GUID IDs needed in aad.manifest.json.

## Testing Your Implementation

1. **Validate Permissions**: Ensure all required permissions are added to both azurebot.bicep and aad.manifest.json
2. **Test Authentication**: Verify sign-in flow works with the new permissions
3. **Test Actions**: Validate each AI action works correctly
4. **Test Cards**: Ensure cards display correctly for the workload data
5. **Error Handling**: Test error scenarios (permissions denied, not found, etc.)

## Common Issues and Solutions

### Permission Issues
- **Problem**: "Access denied" errors
- **Solution**: Check both aad.manifest.json and azurebot.bicep have matching permissions

### Template Replacement Issues
- **Problem**: Placeholder tokens still visible in generated code
- **Solution**: Ensure all template files are processed and placeholders replaced

### Card Display Issues
- **Problem**: Cards show "undefined" or missing data
- **Solution**: Update card helper methods to use correct property names for workload

### Action Not Found
- **Problem**: AI doesn't recognize custom actions
- **Solution**: Ensure actions.json is updated and prompt folder is correctly referenced

## Example: Complete Calendar Implementation

See the `/examples/calendar/` directory (to be created) for a complete working calendar implementation generated from this template.

This shows the exact file structure and code that should result from following this guide for a calendar workload.