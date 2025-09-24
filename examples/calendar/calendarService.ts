import { BaseGraphService } from '../../src/services/baseGraphService';

export interface CalendarItem {
  id: string;
  subject: string;
  start: { dateTime: string; timeZone: string };
  end: { dateTime: string; timeZone: string };
  body?: { content: string; contentType?: string };
  attendees?: Array<{ emailAddress: { address: string; name: string } }>;
  location?: { displayName: string };
  isAllDay?: boolean;
  importance?: string;
  sensitivity?: string;
  showAs?: string;
  responseStatus?: { response: string; time: string };
  organizer?: { emailAddress: { address: string; name: string } };
  createdDateTime: string;
  lastModifiedDateTime?: string;
}

export interface CalendarCreateRequest {
  subject: string;
  start: { dateTime: string; timeZone: string };
  end: { dateTime: string; timeZone: string };
  body?: { content: string; contentType?: string };
  attendees?: Array<{ emailAddress: { address: string; name: string } }>;
  location?: { displayName: string };
  isAllDay?: boolean;
  importance?: string;
}

export interface CalendarUpdateRequest {
  subject?: string;
  start?: { dateTime: string; timeZone: string };
  end?: { dateTime: string; timeZone: string };
  body?: { content: string; contentType?: string };
  attendees?: Array<{ emailAddress: { address: string; name: string } }>;
  location?: { displayName: string };
  isAllDay?: boolean;
  importance?: string;
}

export interface CalendarQueryOptions {
  select?: string[];
  filter?: string;
  orderBy?: string;
  top?: number;
  skip?: number;
  expand?: string[];
}

export const CALENDAR_PERMISSIONS = {
  READ: 'Calendars.Read',
  WRITE: 'Calendars.ReadWrite',
  SHARED: 'Calendars.ReadWrite.Shared'
} as const;

export const CALENDAR_ENDPOINTS = {
  BASE: '/me/events',
  CALENDARS: '/me/calendars',
  CALENDAR_VIEW: '/me/calendarView'
} as const;

export class CalendarService extends BaseGraphService {

  /**
   * Get all calendar events for the current user
   */
  async getEvents(options?: CalendarQueryOptions): Promise<CalendarItem[]> {
    try {
      const endpoint = CALENDAR_ENDPOINTS.BASE;

      const queryParams: Record<string, any> = {};
      if (options?.select) queryParams.$select = options.select.join(',');
      if (options?.filter) queryParams.$filter = options.filter;
      if (options?.orderBy) queryParams.$orderby = options.orderBy;
      if (options?.top) queryParams.$top = options.top;
      if (options?.skip) queryParams.$skip = options.skip;
      if (options?.expand) queryParams.$expand = options.expand.join(',');

      const events = await this.getPaginatedResults<CalendarItem>(endpoint, queryParams);
      return events;
    } catch (error: any) {
      throw this.handleGraphError(error, 'retrieve calendar events');
    }
  }

  /**
   * Create a new calendar event
   */
  async createEvent(data: CalendarCreateRequest): Promise<CalendarItem> {
    try {
      if (!data.subject?.trim()) {
        throw new Error('Event subject is required');
      }
      if (!data.start) {
        throw new Error('Event start time is required');
      }
      if (!data.end) {
        throw new Error('Event end time is required');
      }

      const endpoint = CALENDAR_ENDPOINTS.BASE;
      const result = await this.createResource<CalendarItem>(endpoint, data);
      return result;
    } catch (error: any) {
      throw this.handleGraphError(error, 'create calendar event');
    }
  }

  /**
   * Get a specific calendar event by ID
   */
  async getEventById(id: string, options?: CalendarQueryOptions): Promise<CalendarItem> {
    try {
      if (!id?.trim()) {
        throw new Error('Event ID is required');
      }

      let endpoint = `${CALENDAR_ENDPOINTS.BASE}/${id}`;
      let request = this.graphClient.api(endpoint);

      if (options?.select) request = request.select(options.select);
      if (options?.expand) request = request.expand(options.expand);

      const result = await request.get();
      return result;
    } catch (error: any) {
      throw this.handleGraphError(error, 'retrieve calendar event');
    }
  }

  /**
   * Update a calendar event
   */
  async updateEvent(id: string, updates: CalendarUpdateRequest): Promise<CalendarItem> {
    try {
      if (!id?.trim()) {
        throw new Error('Event ID is required');
      }
      if (!updates || Object.keys(updates).length === 0) {
        throw new Error('Update data is required');
      }

      const endpoint = `${CALENDAR_ENDPOINTS.BASE}/${id}`;
      const result = await this.updateResource<CalendarItem>(endpoint, updates);
      return result;
    } catch (error: any) {
      throw this.handleGraphError(error, 'update calendar event');
    }
  }

  /**
   * Delete a calendar event
   */
  async deleteEvent(id: string): Promise<void> {
    try {
      if (!id?.trim()) {
        throw new Error('Event ID is required');
      }

      const endpoint = `${CALENDAR_ENDPOINTS.BASE}/${id}`;
      await this.deleteResource(endpoint);
    } catch (error: any) {
      throw this.handleGraphError(error, 'delete calendar event');
    }
  }

  /**
   * Get events for current user (alias for getEvents for consistency)
   */
  async getMyEvents(): Promise<CalendarItem[]> {
    return this.getEvents();
  }

  /**
   * Get upcoming events for the next N days
   */
  async getUpcomingEvents(days: number = 7): Promise<CalendarItem[]> {
    try {
      const now = new Date();
      const endDate = new Date(now.getTime() + (days * 24 * 60 * 60 * 1000));

      const filter = `start/dateTime ge '${now.toISOString()}' and start/dateTime le '${endDate.toISOString()}'`;

      return await this.getEvents({
        filter,
        orderBy: 'start/dateTime',
        top: 50
      });
    } catch (error: any) {
      throw this.handleGraphError(error, 'retrieve upcoming events');
    }
  }

  /**
   * Create a meeting with attendees
   */
  async createMeeting(
    subject: string,
    start: { dateTime: string; timeZone: string },
    end: { dateTime: string; timeZone: string },
    attendees: string[],
    body?: string,
    location?: string
  ): Promise<CalendarItem> {
    try {
      const meetingData: CalendarCreateRequest = {
        subject,
        start,
        end,
        attendees: attendees.map(email => ({
          emailAddress: { address: email, name: email }
        }))
      };

      if (body) {
        meetingData.body = { content: body, contentType: 'HTML' };
      }

      if (location) {
        meetingData.location = { displayName: location };
      }

      return await this.createEvent(meetingData);
    } catch (error: any) {
      throw this.handleGraphError(error, 'create meeting');
    }
  }

  /**
   * Get events for today
   */
  async getTodaysEvents(): Promise<CalendarItem[]> {
    try {
      const today = new Date();
      const startOfDay = new Date(today.getFullYear(), today.getMonth(), today.getDate());
      const endOfDay = new Date(today.getFullYear(), today.getMonth(), today.getDate() + 1);

      const filter = `start/dateTime ge '${startOfDay.toISOString()}' and start/dateTime lt '${endOfDay.toISOString()}'`;

      return await this.getEvents({
        filter,
        orderBy: 'start/dateTime'
      });
    } catch (error: any) {
      throw this.handleGraphError(error, 'retrieve today\'s events');
    }
  }

  /**
   * Get calendar view (events in a specific time range)
   */
  async getCalendarView(startTime: string, endTime: string): Promise<CalendarItem[]> {
    try {
      const endpoint = CALENDAR_ENDPOINTS.CALENDAR_VIEW;
      const queryParams = {
        startDateTime: startTime,
        endDateTime: endTime
      };

      return await this.getPaginatedResults<CalendarItem>(endpoint, queryParams);
    } catch (error: any) {
      throw this.handleGraphError(error, 'retrieve calendar view');
    }
  }

  /**
   * Accept a meeting invitation
   */
  async acceptMeeting(eventId: string, comment?: string): Promise<void> {
    try {
      if (!eventId?.trim()) {
        throw new Error('Event ID is required');
      }

      const endpoint = `${CALENDAR_ENDPOINTS.BASE}/${eventId}/accept`;
      const data = comment ? { comment } : {};

      await this.graphClient.api(endpoint).post(data);
    } catch (error: any) {
      throw this.handleGraphError(error, 'accept meeting invitation');
    }
  }

  /**
   * Decline a meeting invitation
   */
  async declineMeeting(eventId: string, comment?: string): Promise<void> {
    try {
      if (!eventId?.trim()) {
        throw new Error('Event ID is required');
      }

      const endpoint = `${CALENDAR_ENDPOINTS.BASE}/${eventId}/decline`;
      const data = comment ? { comment } : {};

      await this.graphClient.api(endpoint).post(data);
    } catch (error: any) {
      throw this.handleGraphError(error, 'decline meeting invitation');
    }
  }

  /**
   * Tentatively accept a meeting invitation
   */
  async tentativelyAcceptMeeting(eventId: string, comment?: string): Promise<void> {
    try {
      if (!eventId?.trim()) {
        throw new Error('Event ID is required');
      }

      const endpoint = `${CALENDAR_ENDPOINTS.BASE}/${eventId}/tentativelyAccept`;
      const data = comment ? { comment } : {};

      await this.graphClient.api(endpoint).post(data);
    } catch (error: any) {
      throw this.handleGraphError(error, 'tentatively accept meeting invitation');
    }
  }
}