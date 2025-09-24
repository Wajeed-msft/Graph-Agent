# Calendar Agent Implementation Example

This directory contains a complete implementation of a Calendar agent using the Graph Agent template. This serves as a reference for how to implement any Microsoft Graph workload using the template system.

## Files Created

### Service Layer
- `calendarService.ts` - Complete Calendar service implementation
  - Implements all CRUD operations for calendar events
  - Includes calendar-specific methods like `createMeeting`, `getUpcomingEvents`, `acceptMeeting`
  - Proper error handling and type safety

### UI Layer
- `calendarCards.ts` - Calendar-specific Adaptive Cards
  - Event list cards with time formatting
  - Detailed event cards with all properties
  - Today's events card with simplified view
  - Event creation confirmation cards

## Key Implementation Details

### 1. Template Replacements Made
- `{WORKLOAD_NAME}` → `Calendar`
- `{WORKLOAD_NAME_LOWER}` → `calendar`
- `{WORKLOAD_ICON}` → `📅`
- `{WORKLOAD_PRIMARY_FIELD}` → `subject`
- `{WORKLOAD_BASE_ENDPOINT}` → `/me/events`

### 2. Calendar-Specific Types
```typescript
interface CalendarItem {
  id: string;
  subject: string;
  start: { dateTime: string; timeZone: string };
  end: { dateTime: string; timeZone: string };
  body?: { content: string; contentType?: string };
  attendees?: Array<{ emailAddress: { address: string; name: string } }>;
  location?: { displayName: string };
  // ... more properties
}
```

### 3. Calendar-Specific Methods
- `getUpcomingEvents(days)` - Get events for next N days
- `createMeeting()` - Create meeting with attendees
- `getTodaysEvents()` - Get today's calendar
- `acceptMeeting()` / `declineMeeting()` - Respond to invites
- `getCalendarView()` - Events in time range

### 4. Card Customizations
- Time formatting for different scenarios (all-day, same-day, multi-day)
- Event status indicators (accepted, declined, tentative)
- Attendee and location display
- Meeting response actions

## How This Was Generated

This implementation demonstrates what Claude Code should generate when given the prompt:
> "Build me a calendar tracking agent"

### Step-by-step Process:
1. **Identified Workload**: Calendar → Graph API events
2. **Applied Template Replacements**: All placeholders replaced with calendar-specific values
3. **Implemented Service Logic**: Added calendar-specific endpoints and methods
4. **Customized Card Display**: Event-specific formatting and metadata
5. **Added Workload Actions**: Meeting creation, responses, upcoming events
6. **Updated Infrastructure**: Calendar permissions in bicep and manifest files

## Using This Implementation

To use this calendar implementation:

1. Copy files to main src directory:
   ```bash
   cp calendarService.ts ../../src/services/
   cp calendarCards.ts ../../src/cards/
   ```

2. Update main index.ts imports:
   ```typescript
   import { CalendarService } from './services/calendarService';
   import { CalendarCards } from './cards/calendarCards';
   ```

3. Update infrastructure files with calendar permissions:
   - `infra/botRegistration/azurebot.bicep`: Set `scopes: 'Calendars.Read,Calendars.ReadWrite,User.Read'`
   - `aad.manifest.json`: Add Calendars.Read and Calendars.ReadWrite permissions

4. Update prompts directory:
   ```bash
   cp -r ../../src/prompts/generic ../../src/prompts/calendar
   # Edit files to replace placeholders with calendar values
   ```

5. Update prompt manager in index.ts:
   ```typescript
   defaultPrompt: 'calendar'
   ```

## Testing the Implementation

The calendar agent should support these user interactions:
- "Show me my calendar events"
- "What's on my schedule today?"
- "Create a meeting for tomorrow at 2 PM"
- "Get my upcoming events for this week"
- "Accept the meeting invitation"
- "Show me details for event ABC123"

This implementation provides a complete, working calendar agent that demonstrates the power and flexibility of the Graph Agent template system.