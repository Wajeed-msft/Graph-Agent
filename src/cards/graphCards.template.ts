import { {WORKLOAD_NAME}Item } from '../types/graphTypes.template';

export class {WORKLOAD_NAME}Cards {

  /**
   * Create a list card displaying multiple {WORKLOAD_NAME_LOWER} items
   */
  static create{WORKLOAD_NAME}sListCard(items: {WORKLOAD_NAME}Item[]): any {
    const card = {
      type: "AdaptiveCard",
      version: "1.4",
      body: [
        {
          type: "TextBlock",
          text: "{WORKLOAD_ICON} Your {WORKLOAD_NAME}s",
          size: "Large",
          weight: "Bolder",
          color: "Accent"
        }
      ]
    };

    if (items.length === 0) {
      card.body.push({
        type: "TextBlock",
        text: "You don't have any {WORKLOAD_NAME_LOWER}s yet. Create one by asking me!",
        wrap: true
      });
      return card;
    }

    const itemElements = items.map(item => ({
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
                  text: `**${this.get{WORKLOAD_NAME}Title(item)}**`,
                  size: "Medium",
                  weight: "Bolder",
                  wrap: true
                },
                {
                  type: "TextBlock",
                  text: this.get{WORKLOAD_NAME}Subtitle(item),
                  size: "Small",
                  color: "Accent",
                  wrap: true
                },
                {
                  type: "TextBlock",
                  text: this.get{WORKLOAD_NAME}Metadata(item),
                  size: "Small",
                  color: "Dark",
                  wrap: true
                }
              ]
            }
          ]
        }
      ]
    }));

    card.body.push({
      type: "Container",
      items: itemElements
    });

    return card;
  }

  /**
   * Create a detailed card for a single {WORKLOAD_NAME_LOWER} item
   */
  static create{WORKLOAD_NAME}DetailCard(item: {WORKLOAD_NAME}Item): any {
    return {
      type: "AdaptiveCard",
      version: "1.4",
      body: [
        {
          type: "TextBlock",
          text: "{WORKLOAD_ICON} {WORKLOAD_NAME} Details",
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
              text: `**${this.get{WORKLOAD_NAME}Title(item)}**`,
              size: "Large",
              weight: "Bolder",
              wrap: true
            },
            {
              type: "TextBlock",
              text: this.get{WORKLOAD_NAME}Description(item),
              wrap: true,
              spacing: "Medium"
            },
            {
              type: "FactSet",
              facts: this.get{WORKLOAD_NAME}Facts(item)
            }
            // TODO: Add workload-specific details
            // Example for Calendar: attendees, location, meeting link
            // Example for Teams: member count, channels, privacy
            // Example for SharePoint: file size, last modified, permissions
          ]
        }
      ]
    };
  }

  /**
   * Create a summary card for user-specific {WORKLOAD_NAME_LOWER} items
   */
  static createMy{WORKLOAD_NAME}sCard(items: {WORKLOAD_NAME}Item[]): any {
    const card = {
      type: "AdaptiveCard",
      version: "1.4",
      body: [
        {
          type: "TextBlock",
          text: "👤 My {WORKLOAD_NAME}s",
          size: "Large",
          weight: "Bolder",
          color: "Accent"
        }
      ]
    };

    if (items.length === 0) {
      card.body.push({
        type: "TextBlock",
        text: "You have no {WORKLOAD_NAME_LOWER}s assigned to you.",
        wrap: true
      });
      return card;
    }

    // TODO: Customize based on workload
    // Example for Calendar: upcoming events, today's schedule
    // Example for Teams: teams you're in, recent activity
    // Example for SharePoint: recent files, shared documents

    const itemElements = items.map(item => {
      const itemContainer = {
        type: "Container",
        style: "emphasis",
        spacing: "Small",
        items: [
          {
            type: "TextBlock",
            text: this.get{WORKLOAD_NAME}Title(item),
            weight: "Bolder",
            wrap: true
          },
          {
            type: "TextBlock",
            text: this.get{WORKLOAD_NAME}Status(item),
            size: "Small",
            color: this.get{WORKLOAD_NAME}StatusColor(item)
          },
          {
            type: "TextBlock",
            text: this.get{WORKLOAD_NAME}Metadata(item),
            size: "Small",
            color: "Dark"
          }
        ]
      };

      return itemContainer;
    });

    card.body.push({
      type: "Container",
      items: itemElements
    });

    return card;
  }

  /**
   * Create a creation confirmation card
   */
  static create{WORKLOAD_NAME}CreatedCard(item: {WORKLOAD_NAME}Item): any {
    return {
      type: "AdaptiveCard",
      version: "1.4",
      body: [
        {
          type: "TextBlock",
          text: "✅ {WORKLOAD_NAME} Created Successfully!",
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
              text: `**${this.get{WORKLOAD_NAME}Title(item)}**`,
              size: "Medium",
              weight: "Bolder",
              wrap: true
            },
            {
              type: "TextBlock",
              text: this.get{WORKLOAD_NAME}Subtitle(item),
              size: "Small",
              color: "Accent",
              wrap: true
            },
            {
              type: "TextBlock",
              text: `ID: ${item.id}`,
              size: "Small",
              color: "Dark",
              wrap: true
            }
          ]
        }
      ]
    };
  }

  // TODO: Implement these helper methods based on workload
  // These should be customized for each specific workload

  /**
   * Get the main title/name of the item
   */
  private static get{WORKLOAD_NAME}Title(item: {WORKLOAD_NAME}Item): string {
    // TODO: Return workload-specific title field
    // Example for Calendar: item.subject
    // Example for Teams: item.displayName
    // Example for SharePoint: item.name
    return item.title || item.name || item.displayName || item.subject || 'Untitled';
  }

  /**
   * Get subtitle/secondary information
   */
  private static get{WORKLOAD_NAME}Subtitle(item: {WORKLOAD_NAME}Item): string {
    // TODO: Return workload-specific subtitle
    // Example for Calendar: date/time range
    // Example for Teams: team description
    // Example for SharePoint: file type or folder path
    return item.id ? `ID: ${item.id.substring(0, 8)}...` : '';
  }

  /**
   * Get metadata information (dates, counts, etc.)
   */
  private static get{WORKLOAD_NAME}Metadata(item: {WORKLOAD_NAME}Item): string {
    // TODO: Return workload-specific metadata
    // Example for Calendar: "Today at 2:00 PM"
    // Example for Teams: "5 members • 3 channels"
    // Example for SharePoint: "Modified 2 days ago • 1.2 MB"
    const createdDate = item.createdDateTime || item.created;
    return createdDate ? `Created: ${new Date(createdDate).toLocaleDateString()}` : '';
  }

  /**
   * Get status information
   */
  private static get{WORKLOAD_NAME}Status(item: {WORKLOAD_NAME}Item): string {
    // TODO: Return workload-specific status
    // Example for Calendar: "Upcoming", "In Progress", "Completed"
    // Example for Teams: "Active", "Archived"
    // Example for SharePoint: "Shared", "Private"
    return item.status || 'Active';
  }

  /**
   * Get status color for display
   */
  private static get{WORKLOAD_NAME}StatusColor(item: {WORKLOAD_NAME}Item): string {
    const status = this.get{WORKLOAD_NAME}Status(item);
    // TODO: Customize based on workload status values
    switch (status.toLowerCase()) {
      case 'completed':
      case 'done':
      case 'success':
        return 'Good';
      case 'in progress':
      case 'active':
      case 'ongoing':
        return 'Warning';
      case 'cancelled':
      case 'failed':
      case 'error':
        return 'Attention';
      default:
        return 'Default';
    }
  }

  /**
   * Get description or detailed information
   */
  private static get{WORKLOAD_NAME}Description(item: {WORKLOAD_NAME}Item): string {
    // TODO: Return workload-specific description
    // Example for Calendar: item.body?.content
    // Example for Teams: item.description
    // Example for SharePoint: item.description or file content preview
    return item.description || item.body?.content || 'No description available';
  }

  /**
   * Get fact set for detailed view
   */
  private static get{WORKLOAD_NAME}Facts(item: {WORKLOAD_NAME}Item): any[] {
    // TODO: Return workload-specific facts
    const facts = [];

    if (item.id) {
      facts.push({ title: 'ID', value: item.id });
    }

    // Add workload-specific facts
    // Example for Calendar:
    // if (item.start) facts.push({ title: 'Start', value: new Date(item.start.dateTime).toLocaleString() });
    // if (item.end) facts.push({ title: 'End', value: new Date(item.end.dateTime).toLocaleString() });

    // Example for Teams:
    // if (item.memberCount) facts.push({ title: 'Members', value: item.memberCount.toString() });
    // if (item.visibility) facts.push({ title: 'Privacy', value: item.visibility });

    // Example for SharePoint:
    // if (item.size) facts.push({ title: 'Size', value: this.formatFileSize(item.size) });
    // if (item.lastModifiedDateTime) facts.push({ title: 'Modified', value: new Date(item.lastModifiedDateTime).toLocaleString() });

    const createdDate = item.createdDateTime || item.created;
    if (createdDate) {
      facts.push({ title: 'Created', value: new Date(createdDate).toLocaleString() });
    }

    return facts;
  }
}