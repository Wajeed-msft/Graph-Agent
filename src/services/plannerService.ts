import { Client } from '@microsoft/microsoft-graph-client';
import 'isomorphic-fetch';

export interface PlannerPlan {
  id: string;
  title: string;
  createdDateTime: string;
  owner?: string;
  createdBy?: {
    user?: {
      displayName?: string;
      id?: string;
    };
  };
  container?: {
    containerId?: string;
    type?: string;
    url?: string;
  };
  groupName?: string;
  groupDescription?: string;
}

export interface PlannerTask {
  id: string;
  title: string;
  planId: string;
  bucketId: string;
  assigneePriority?: string;
  assignments?: { [key: string]: any };
  percentComplete: number;
  createdDateTime: string;
}

export interface PlannerBucket {
  id: string;
  name: string;
  planId: string;
  orderHint: string;
}

export class PlannerService {
  private graphClient: Client;

  constructor(private accessToken: string) {
    // Log the token for debugging purposes
    console.log('🔑 Microsoft Graph Access Token:', this.accessToken);
    console.log('🔑 Token length:', this.accessToken.length);
    console.log('🔑 Token starts with:', this.accessToken.substring(0, 20) + '...');
    
    // Initialize the Graph client with a custom auth provider
    // The callback signature expects (error, accessToken)
    const authProvider = (done: (error: any, accessToken: string | null) => void) => {
      try {
        console.log('🔐 AuthProvider called - providing token for Graph API call');
        done(null, this.accessToken);
      } catch (error) {
        console.error('🚨 AuthProvider error:', error);
        done(error, null);
      }
    };

    this.graphClient = Client.init({
      authProvider: authProvider
    });
  }

  async getPlans(): Promise<PlannerPlan[]> {
    try {
      const response = await this.graphClient.api('/me/planner/plans')
        .select('id,title,createdDateTime,owner,createdBy,container')
        .get();
      
      const plans = response.value || [];
      
      // Try to enhance plans with group information
      const enhancedPlans = await Promise.allSettled(
        plans.map(async (plan: any) => {
          try {
            // If the plan has a container with a group, try to get group info
            if (plan.container?.containerId) {
              const groupInfo = await this.graphClient
                .api(`/groups/${plan.container.containerId}`)
                .select('displayName,description')
                .get();
              
              return {
                ...plan,
                groupName: groupInfo.displayName,
                groupDescription: groupInfo.description
              };
            }
            return plan;
          } catch {
            // If we can't get group info, return plan as-is
            return plan;
          }
        })
      );
      
      return enhancedPlans
        .filter((result): result is PromiseFulfilledResult<any> => result.status === 'fulfilled')
        .map(result => result.value);
        
    } catch (error: any) {
      console.error('Error fetching plans:', error);
      if (error.code === 'Forbidden') {
        throw new Error('Access denied. Please ensure you have the required permissions to access Planner.');
      } else if (error.code === 'NotFound') {
        throw new Error('Planner service not found. Please check if Planner is available in your organization.');
      }
      throw new Error(`Failed to retrieve plans: ${error.message || 'Unknown error'}`);
    }
  }

  async createPlan(title: string, groupId: string): Promise<PlannerPlan> {
    try {
      if (!title?.trim()) {
        throw new Error('Plan title is required');
      }
      if (!groupId?.trim()) {
        throw new Error('Group ID is required');
      }

      const planData = {
        title: title.trim(),
        container: {
          url: `https://graph.microsoft.com/v1.0/groups/${groupId}`
        }
      };

      const response = await this.graphClient.api('/planner/plans').post(planData);
      return response;
    } catch (error: any) {
      console.error('Error creating plan:', error);
      if (error.code === 'Forbidden') {
        throw new Error('Access denied. You must be a member of the group to create plans.');
      } else if (error.code === 'BadRequest') {
        throw new Error('Invalid plan data. Please check the title and group membership.');
      }
      throw new Error(`Failed to create plan: ${error.message || 'Unknown error'}`);
    }
  }

  async getPlanTasks(planId: string): Promise<PlannerTask[]> {
    try {
      if (!planId?.trim()) {
        throw new Error('Plan ID is required');
      }
      const response = await this.graphClient.api(`/planner/plans/${planId}/tasks`).get();
      return response.value || [];
    } catch (error: any) {
      console.error('Error fetching plan tasks:', error);
      if (error.code === 'NotFound') {
        throw new Error('Plan not found. Please check the plan ID.');
      } else if (error.code === 'Forbidden') {
        throw new Error('Access denied. You may not have permission to view this plan.');
      }
      throw new Error(`Failed to retrieve tasks: ${error.message || 'Unknown error'}`);
    }
  }

  async getPlanBuckets(planId: string): Promise<PlannerBucket[]> {
    try {
      if (!planId?.trim()) {
        throw new Error('Plan ID is required');
      }
      const response = await this.graphClient.api(`/planner/plans/${planId}/buckets`).get();
      return response.value || [];
    } catch (error: any) {
      console.error('Error fetching plan buckets:', error);
      if (error.code === 'NotFound') {
        throw new Error('Plan not found. Please check the plan ID.');
      } else if (error.code === 'Forbidden') {
        throw new Error('Access denied. You may not have permission to view this plan.');
      }
      throw new Error(`Failed to retrieve buckets: ${error.message || 'Unknown error'}`);
    }
  }

  async createTask(planId: string, bucketId: string, title: string, assigneeId?: string): Promise<PlannerTask> {
    try {
      if (!planId?.trim()) {
        throw new Error('Plan ID is required');
      }
      if (!bucketId?.trim()) {
        throw new Error('Bucket ID is required');
      }
      if (!title?.trim()) {
        throw new Error('Task title is required');
      }

      const taskData: any = {
        title: title.trim(),
        planId: planId.trim(),
        bucketId: bucketId.trim()
      };

      if (assigneeId?.trim()) {
        taskData.assignments = {
          [assigneeId]: {
            '@odata.type': 'microsoft.graph.plannerAssignment',
            orderHint: ' !'
          }
        };
      }

      const response = await this.graphClient.api('/planner/tasks').post(taskData);
      return response;
    } catch (error: any) {
      console.error('Error creating task:', error);
      if (error.code === 'Forbidden') {
        throw new Error('Access denied. You may not have permission to create tasks in this plan.');
      } else if (error.code === 'BadRequest') {
        throw new Error('Invalid task data. Please check the plan and bucket IDs.');
      }
      throw new Error(`Failed to create task: ${error.message || 'Unknown error'}`);
    }
  }

  async createTaskWithoutBucket(planId: string, title: string, assigneeId?: string): Promise<PlannerTask> {
    try {
      if (!planId?.trim()) {
        throw new Error('Plan ID is required');
      }
      if (!title?.trim()) {
        throw new Error('Task title is required');
      }

      // First try to get any bucket from the plan
      let bucketId: string | undefined;
      try {
        const buckets = await this.getPlanBuckets(planId);
        if (buckets.length > 0) {
          bucketId = buckets[0].id;
        }
      } catch {
        // If we can't get buckets, we'll need to create a default bucket or handle this differently
      }

      if (!bucketId) {
        // If no bucket found, try to create a default one
        try {
          const defaultBucket = await this.createDefaultBucket(planId);
          bucketId = defaultBucket.id;
        } catch {
          throw new Error('Cannot create task: No buckets available and cannot create default bucket');
        }
      }

      const taskData: any = {
        title: title.trim(),
        planId: planId.trim(),
        bucketId: bucketId
      };

      if (assigneeId?.trim()) {
        taskData.assignments = {
          [assigneeId]: {
            '@odata.type': 'microsoft.graph.plannerAssignment',
            orderHint: ' !'
          }
        };
      }

      const response = await this.graphClient.api('/planner/tasks').post(taskData);
      return response;
    } catch (error: any) {
      console.error('Error creating task without bucket:', error);
      if (error.code === 'Forbidden') {
        throw new Error('Access denied. You may not have permission to create tasks in this plan.');
      } else if (error.code === 'BadRequest') {
        throw new Error('Invalid task data. Please check the plan ID.');
      }
      throw new Error(`Failed to create task: ${error.message || 'Unknown error'}`);
    }
  }

  private async createDefaultBucket(planId: string): Promise<PlannerBucket> {
    try {
      const bucketData = {
        name: 'To Do',
        planId: planId,
        orderHint: ' !'
      };

      const response = await this.graphClient.api('/planner/buckets').post(bucketData);
      return response;
    } catch (error: any) {
      console.error('Error creating default bucket:', error);
      throw new Error(`Failed to create default bucket: ${error.message || 'Unknown error'}`);
    }
  }

  async updateTask(taskId: string, updates: Partial<PlannerTask>): Promise<PlannerTask> {
    try {
      if (!taskId?.trim()) {
        throw new Error('Task ID is required');
      }
      if (!updates || Object.keys(updates).length === 0) {
        throw new Error('Update data is required');
      }

      const response = await this.graphClient.api(`/planner/tasks/${taskId}`).patch(updates);
      return response;
    } catch (error: any) {
      console.error('Error updating task:', error);
      if (error.code === 'NotFound') {
        throw new Error('Task not found. Please check the task ID.');
      } else if (error.code === 'Forbidden') {
        throw new Error('Access denied. You may not have permission to update this task.');
      } else if (error.code === 'PreconditionFailed') {
        throw new Error('Task has been modified by another user. Please refresh and try again.');
      }
      throw new Error(`Failed to update task: ${error.message || 'Unknown error'}`);
    }
  }

  async getMyTasks(): Promise<PlannerTask[]> {
    try {
      const response = await this.graphClient.api('/me/planner/tasks').get();
      return response.value || [];
    } catch (error: any) {
      console.error('Error fetching my tasks:', error);
      if (error.code === 'Forbidden') {
        throw new Error('Access denied. Please ensure you have the required permissions to access your tasks.');
      }
      throw new Error(`Failed to retrieve your tasks: ${error.message || 'Unknown error'}`);
    }
  }

  async getUserGroups(): Promise<any[]> {
    try {
      const response = await this.graphClient.api('/me/memberOf').get();
      return response.value?.filter((group: any) => group['@odata.type'] === '#microsoft.graph.group') || [];
    } catch (error: any) {
      console.error('Error fetching user groups:', error);
      if (error.code === 'Forbidden') {
        throw new Error('Access denied. Please ensure you have the required permissions to access group information.');
      }
      throw new Error(`Failed to retrieve groups: ${error.message || 'Unknown error'}`);
    }
  }

  async getCurrentUser(): Promise<any> {
    try {
      const response = await this.graphClient.api('/me').get();
      return response;
    } catch (error: any) {
      console.error('Error fetching current user:', error);
      if (error.code === 'Forbidden') {
        throw new Error('Access denied. Please ensure you have the required permissions to access user information.');
      }
      throw new Error(`Failed to retrieve current user: ${error.message || 'Unknown error'}`);
    }
  }
}