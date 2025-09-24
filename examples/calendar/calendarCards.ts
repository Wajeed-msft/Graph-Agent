import { CalendarItem } from './calendarService';

export class CalendarCards {

  /**
   * Create a list card displaying multiple calendar events
   */
  static createEventsListCard(events: CalendarItem[]): any {
    const card = {
      type: "AdaptiveCard",
      version: "1.4",
      body: [
        {
          type: "TextBlock",
          text: "📅 Your Calendar Events",
          size: "Large",
          weight: "Bolder",
          color: "Accent"
        }
      ]
    };

    if (events.length === 0) {
      card.body.push({
        type: "TextBlock",
        text: "You don't have any calendar events yet. Create one by asking me!",
        wrap: true
      });
      return card;
    }

    const eventElements = events.map(event => ({
      type: "Container",
      style: "emphasis",
      spacing: "Medium",
      items: [
        {
          type: "ColumnSet",
          columns: [
            {
              type: "Column",
              width: "stretch",
              items: [
                {
                  type: "TextBlock",
                  text: `**${this.getEventTitle(event)}**`,
                  size: "Medium",
                  weight: "Bolder",
                  wrap: true
                },
                {
                  type: "TextBlock",
                  text: this.getEventTimeDisplay(event),
                  size: "Small",
                  color: "Accent",
                  wrap: true
                },
                {
                  type: "TextBlock",
                  text: this.getEventMetadata(event),
                  size: "Small",
                  color: "Dark",
                  wrap: true
                }
              ]
            },
            {
              type: "Column",
              width: "auto",
              items: [
                {
                  type: "TextBlock",
                  text: this.getEventStatusIcon(event),
                  size: "Medium"
                }
              ]
            }
          ]
        }
      ]
    }));

    card.body.push({
      type: "Container",
      items: eventElements
    });

    return card;
  }

  /**
   * Create a detailed card for a single calendar event
   */
  static createEventDetailCard(event: CalendarItem): any {
    return {
      type: "AdaptiveCard",
      version: "1.4",
      body: [
        {
          type: "TextBlock",
          text: "📅 Event Details",
          size: "Large",
          weight: "Bolder",
          color: "Accent"
        },
        {
          type: "Container",
          style: "emphasis",
          spacing: "Medium",
          items: [
            {
              type: "TextBlock",
              text: `**${this.getEventTitle(event)}**`,
              size: "Large",
              weight: "Bolder",
              wrap: true
            },
            {
              type: "TextBlock",
              text: this.getEventDescription(event),
              wrap: true,
              spacing: "Medium"
            },
            {
              type: "FactSet",
              facts: this.getEventFacts(event)
            }
          ]
        }
      ]
    };
  }

  /**
   * Create a summary card for user's calendar events
   */
  static createMyEventsCard(events: CalendarItem[]): any {
    const card = {
      type: "AdaptiveCard",
      version: "1.4",
      body: [
        {
          type: "TextBlock",
          text: "👤 My Calendar Events",
          size: "Large",
          weight: "Bolder",
          color: "Accent"
        }
      ]
    };

    if (events.length === 0) {
      card.body.push({
        type: "TextBlock",
        text: "You have no calendar events.",
        wrap: true
      });
      return card;
    }

    const eventElements = events.map(event => {
      const eventContainer = {
        type: "Container",
        style: "emphasis",
        spacing: "Small",
        items: [
          {
            type: "TextBlock",
            text: this.getEventTitle(event),
            weight: "Bolder",
            wrap: true
          },
          {
            type: "TextBlock",
            text: this.getEventStatus(event),
            size: "Small",
            color: this.getEventStatusColor(event)
          },
          {
            type: "TextBlock",
            text: this.getEventTimeDisplay(event),
            size: "Small",
            color: "Dark"
          }
        ]
      };

      return eventContainer;
    });

    card.body.push({
      type: "Container",
      items: eventElements
    });

    return card;
  }

  /**
   * Create an event creation confirmation card
   */
  static createEventCreatedCard(event: CalendarItem): any {
    return {
      type: "AdaptiveCard",
      version: "1.4",
      body: [
        {
          type: "TextBlock",
          text: "✅ Event Created Successfully!",
          size: "Large",
          weight: "Bolder",
          color: "Good"
        },
        {
          type: "Container",
          style: "emphasis",
          spacing: "Medium",
          items: [
            {
              type: "TextBlock",
              text: `**${this.getEventTitle(event)}**`,
              size: "Medium",
              weight: "Bolder",
              wrap: true
            },
            {
              type: "TextBlock",
              text: this.getEventTimeDisplay(event),
              size: "Small",
              color: "Accent",
              wrap: true
            },
            {
              type: "TextBlock",
              text: `Event ID: ${event.id}`,
              size: "Small",
              color: "Dark",
              wrap: true
            }
          ]
        }
      ]
    };
  }

  /**
   * Create a card for today's events
   */
  static createTodaysEventsCard(events: CalendarItem[]): any {
    const card = {
      type: "AdaptiveCard",
      version: "1.4",
      body: [
        {
          type: "TextBlock",
          text: "📅 Today's Events",
          size: "Large",
          weight: "Bolder",
          color: "Accent"
        }
      ]
    };

    if (events.length === 0) {
      card.body.push({
        type: "TextBlock",
        text: "No events scheduled for today.",
        wrap: true
      });
      return card;
    }

    const eventElements = events.map(event => ({
      type: "Container",
      style: "emphasis",
      spacing: "Small",
      items: [
        {
          type: "ColumnSet",
          columns: [
            {
              type: "Column",
              width: "auto",
              items: [
                {
                  type: "TextBlock",
                  text: this.getEventTimeOnly(event),
                  size: "Medium",
                  weight: "Bolder",
                  color: "Accent"
                }
              ]
            },
            {
              type: "Column",
              width: "stretch",
              items: [
                {
                  type: "TextBlock",
                  text: this.getEventTitle(event),
                  weight: "Bolder",
                  wrap: true
                },
                {
                  type: "TextBlock",
                  text: this.getEventLocation(event),
                  size: "Small",
                  color: "Dark"
                }
              ]
            }
          ]
        }
      ]
    }));

    card.body.push({
      type: "Container",
      items: eventElements
    });

    return card;
  }

  // Helper methods for extracting and formatting event data

  private static getEventTitle(event: CalendarItem): string {
    return event.subject || 'Untitled Event';
  }

  private static getEventTimeDisplay(event: CalendarItem): string {
    const start = new Date(event.start.dateTime);
    const end = new Date(event.end.dateTime);

    if (event.isAllDay) {
      return `All day • ${start.toLocaleDateString()}`;
    }

    const sameDay = start.toDateString() === end.toDateString();
    if (sameDay) {
      return `${start.toLocaleDateString()} • ${start.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })} - ${end.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}`;
    } else {
      return `${start.toLocaleString([], { month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' })} - ${end.toLocaleString([], { month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' })}`;
    }
  }

  private static getEventTimeOnly(event: CalendarItem): string {
    if (event.isAllDay) {
      return 'All Day';
    }

    const start = new Date(event.start.dateTime);
    return start.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
  }

  private static getEventMetadata(event: CalendarItem): string {
    const parts = [];

    if (event.location?.displayName) {
      parts.push(`📍 ${event.location.displayName}`);
    }

    if (event.attendees && event.attendees.length > 0) {
      parts.push(`👥 ${event.attendees.length} attendees`);
    }

    if (event.organizer?.emailAddress?.name) {
      parts.push(`Organizer: ${event.organizer.emailAddress.name}`);
    }

    return parts.join(' • ') || `Created: ${new Date(event.createdDateTime).toLocaleDateString()}`;
  }

  private static getEventStatus(event: CalendarItem): string {
    if (event.responseStatus) {
      switch (event.responseStatus.response.toLowerCase()) {
        case 'accepted': return 'Accepted';
        case 'declined': return 'Declined';
        case 'tentativelyaccepted': return 'Tentative';
        case 'notresponded': return 'No Response';
        default: return 'Pending';
      }
    }

    const now = new Date();
    const eventStart = new Date(event.start.dateTime);
    const eventEnd = new Date(event.end.dateTime);

    if (eventEnd < now) return 'Past';
    if (eventStart <= now && eventEnd >= now) return 'In Progress';
    return 'Upcoming';
  }

  private static getEventStatusColor(event: CalendarItem): string {
    const status = this.getEventStatus(event);
    switch (status.toLowerCase()) {
      case 'accepted':
      case 'in progress':
        return 'Good';
      case 'declined':
      case 'past':
        return 'Attention';
      case 'tentative':
        return 'Warning';
      default:
        return 'Default';
    }
  }

  private static getEventStatusIcon(event: CalendarItem): string {
    const status = this.getEventStatus(event);
    switch (status.toLowerCase()) {
      case 'accepted': return '✅';
      case 'declined': return '❌';
      case 'tentative': return '⚠️';
      case 'in progress': return '🟢';
      case 'past': return '⏰';
      default: return '📅';
    }
  }

  private static getEventDescription(event: CalendarItem): string {
    return event.body?.content || 'No description available';
  }

  private static getEventLocation(event: CalendarItem): string {
    return event.location?.displayName || '';
  }

  private static getEventFacts(event: CalendarItem): any[] {
    const facts = [];

    if (event.id) {
      facts.push({ title: 'Event ID', value: event.id });
    }

    facts.push({ title: 'Start Time', value: new Date(event.start.dateTime).toLocaleString() });
    facts.push({ title: 'End Time', value: new Date(event.end.dateTime).toLocaleString() });

    if (event.location?.displayName) {
      facts.push({ title: 'Location', value: event.location.displayName });
    }

    if (event.attendees && event.attendees.length > 0) {
      facts.push({ title: 'Attendees', value: event.attendees.length.toString() });
    }

    if (event.organizer?.emailAddress?.name) {
      facts.push({ title: 'Organizer', value: event.organizer.emailAddress.name });
    }

    if (event.importance) {
      facts.push({ title: 'Importance', value: event.importance });
    }

    facts.push({ title: 'Created', value: new Date(event.createdDateTime).toLocaleString() });

    if (event.lastModifiedDateTime) {
      facts.push({ title: 'Modified', value: new Date(event.lastModifiedDateTime).toLocaleString() });
    }

    return facts;
  }
}