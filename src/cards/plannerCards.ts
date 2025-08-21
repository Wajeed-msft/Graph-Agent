import { PlannerPlan, PlannerTask, PlannerBucket } from '../services/plannerService';

export class PlannerCards {
  static createPlansListCard(plans: PlannerPlan[]): any {
    const card = {
      type: "AdaptiveCard",
      version: "1.4",
      body: [
        {
          type: "TextBlock",
          text: "📋 Your Plans",
          size: "Large",
          weight: "Bolder",
          color: "Accent"
        }
      ]
    };

    if (plans.length === 0) {
      card.body.push({
        type: "TextBlock",
        text: "You don't have any plans yet. Create a plan by asking me!",
        wrap: true
      });
      return card;
    }

    const planItems = plans.map(plan => ({
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
                  text: `**${plan.title}**`,
                  size: "Medium",
                  weight: "Bolder",
                  wrap: true
                },
                {
                  type: "TextBlock",
                  text: plan.groupName ? `Group: ${plan.groupName}` : `Plan ID: ${plan.id.substring(0, 8)}...`,
                  size: "Small",
                  color: "Accent",
                  wrap: true
                },
                {
                  type: "TextBlock",
                  text: `Created: ${new Date(plan.createdDateTime).toLocaleDateString()}`,
                  size: "Small",
                  color: "Dark",
                  wrap: true
                },
                {
                  type: "TextBlock",
                  text: plan.createdBy?.user?.displayName ? `Created by: ${plan.createdBy.user.displayName}` : '',
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
      items: planItems
    });

    return card;
  }

  static createTasksBoardCard(tasks: PlannerTask[], buckets: PlannerBucket[], planTitle: string): any {
    const card = {
      type: "AdaptiveCard",
      version: "1.4",
      body: [
        {
          type: "TextBlock",
          text: `📋 ${planTitle} - Tasks`,
          size: "Large",
          weight: "Bolder",
          color: "Accent"
        }
      ]
    };

    if (tasks.length === 0) {
      card.body.push({
        type: "TextBlock",
        text: "No tasks found in this plan.",
        wrap: true
      });
      return card;
    }

    if (buckets.length === 0) {
      // No buckets available, show tasks in a simple list
      const taskItems = tasks.map(task => ({
        type: "Container",
        style: "emphasis",
        spacing: "Small",
        items: [
          {
            type: "TextBlock",
            text: task.title,
            weight: "Bolder",
            wrap: true
          },
          ...(task.percentComplete > 0 ? [{
            type: "TextBlock",
            text: `${task.percentComplete}% Complete`,
            size: "Small",
            color: "Good"
          }] : []),
          {
            type: "TextBlock",
            text: `Created: ${new Date(task.createdDateTime).toLocaleDateString()}`,
            size: "Small",
            color: "Dark"
          }
        ]
      }));

      if (taskItems.length > 0) {
        card.body.push({
          type: "Container",
          items: taskItems
        });
      }
    } else {
      // Show tasks organized by buckets
      buckets.forEach(bucket => {
        const bucketTasks = tasks.filter(task => task.bucketId === bucket.id);
        
        if (bucketTasks.length > 0) {
          const bucketContainer = {
            type: "Container",
            style: "emphasis",
            spacing: "Medium",
            items: [
              {
                type: "TextBlock",
                text: `**${bucket.name}**`,
                size: "Medium",
                weight: "Bolder",
                color: "Attention"
              }
            ]
          };

          bucketTasks.forEach(task => {
            const taskItem = {
              type: "Container",
              style: "default",
              spacing: "Small",
              items: [
                {
                  type: "TextBlock",
                  text: task.title,
                  weight: "Bolder",
                  wrap: true
                }
              ]
            };

            if (task.percentComplete > 0) {
              (taskItem.items as any[]).push({
                type: "TextBlock",
                text: `${task.percentComplete}% Complete`,
                size: "Small",
                color: "Good"
              });
            }

            (taskItem.items as any[]).push({
              type: "TextBlock",
              text: `Created: ${new Date(task.createdDateTime).toLocaleDateString()}`,
              size: "Small",
              color: "Dark"
            });

            (bucketContainer.items as any[]).push(taskItem);
          });

          card.body.push(bucketContainer);
        }
      });
    }

    return card;
  }

  static createMyTasksCard(tasks: PlannerTask[]): any {
    const card = {
      type: "AdaptiveCard",
      version: "1.4",
      body: [
        {
          type: "TextBlock",
          text: "👤 My Assigned Tasks",
          size: "Large",
          weight: "Bolder",
          color: "Accent"
        }
      ]
    };

    if (tasks.length === 0) {
      card.body.push({
        type: "TextBlock",
        text: "You have no assigned tasks.",
        wrap: true
      });
      return card;
    }

    const taskItems = tasks.map(task => {
      const taskContainer = {
        type: "Container",
        style: "emphasis",
        spacing: "Small",
        items: [
          {
            type: "TextBlock",
            text: task.title,
            weight: "Bolder",
            wrap: true
          }
        ]
      };

      if (task.percentComplete > 0) {
        (taskContainer.items as any[]).push({
          type: "TextBlock",
          text: `${task.percentComplete}% Complete`,
          size: "Small",
          color: "Good"
        });
      }

      (taskContainer.items as any[]).push({
        type: "TextBlock",
        text: `Created: ${new Date(task.createdDateTime).toLocaleDateString()}`,
        size: "Small",
        color: "Dark"
      });

      return taskContainer;
    });

    card.body.push({
      type: "Container",
      items: taskItems
    });

    return card;
  }
}