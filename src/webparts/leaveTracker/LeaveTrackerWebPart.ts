import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient } from '@microsoft/sp-http';

import styles from './LeaveTrackerWebPart.module.scss';

export interface ILeaveTrackerWebPartProps {
  description: string;
}

// LeaveTracker Admin List
interface IAdmin {
  Title: string;
  AdminEmail: string;
  Role: string;
  IsActive: boolean;
}

// LeaveTracker Team Members List
interface ITeamMember {
  Title: string;
  EmployeeName: { Title: string, Email: string };
  EmployeeEmail: string;
  LeaveType: string;
  StartDate: string;
  EndDate: string;
  NumberOfDays: number;
  Reason: string;
  ApproveDate?: string;
  RejectionReason?: string;
  RequestDate: string;
  Status: string;
}

// Government Holidays List
interface IGovernmentHoliday {
  Title: string;
  HolidayDate: string;
  HolidayType: string;
  Description: string;
  IsActive: boolean;
}

export default class LeaveTrackerWebPart extends BaseClientSideWebPart<ILeaveTrackerWebPartProps> {
  private leaveTypes: string[] = [];
  private cachedAdmins: IAdmin[] = [];
  private cachedTeamMembers: ITeamMember[] = [];
  private cachedHolidays: IGovernmentHoliday[] = [];
  private isAdmin: boolean = false;

  protected async onInit(): Promise<void> {
    // Load all data once during initialization
    await Promise.all([
      this.loadAdminList(),
      this.loadTeamMembersList(),
      this.loadHolidaysList(),
      this.loadLeaveTypeChoices()
    ]);
    
    this.isAdmin = await this.checkUserAdmin();
    return super.onInit();
  }

  private async loadAdminList(): Promise<void> {
    const listName = "LeaveTracker Admin List";
    try {
      const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
      const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      const data = await response.json();
      this.cachedAdmins = data.value;
      console.log("Admin List Data:", this.cachedAdmins);
    } catch (error) {
      console.error("Error loading Admin List:", error);
      this.cachedAdmins = [];
    }
  }

  private async loadTeamMembersList(): Promise<void> {
    const listName = "LeaveTracker team members list";
    try {
      const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=*,EmployeeName/Title,EmployeeName/EMail&$expand=EmployeeName`;
      const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      const data = await response.json();
      this.cachedTeamMembers = data.value;
      console.log("Team Members List Data:", this.cachedTeamMembers);
    } catch (error) {
      console.error("Error loading Team Members List:", error);
      this.cachedTeamMembers = [];
    }
  }

  private async loadHolidaysList(): Promise<void> {
    const listName = "Government Holidays";
    try {
      const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
      const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      const data = await response.json();
      this.cachedHolidays = data.value;
      console.log("Holidays List Data:", this.cachedHolidays);
    } catch (error) {
      console.error("Error loading Holidays List:", error);
      this.cachedHolidays = [];
    }
  }

  private async loadLeaveTypeChoices(): Promise<void> {
    const listName = "LeaveTracker team members list";
    try {
      const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/fields?$filter=InternalName eq 'LeaveType'`;
      const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      const data = await response.json();

      if (data.value && data.value.length > 0) {
        this.leaveTypes = data.value[0].Choices || [];
      }
    } catch (error) {
      console.error("Error fetching LeaveType choices:", error);
      this.leaveTypes = [];
    }
  }

  private async checkUserAdmin(): Promise<boolean> {
    // const userEmail = this.context.pageContext.user.email;
    const userEmail = "jacob.yeldhos@aciesinnovations.com";
    
    try {
      // Use cached data instead of making new API call
      return this.cachedAdmins.some((admin: IAdmin) => 
        admin.AdminEmail && admin.AdminEmail.toLowerCase() === userEmail.toLowerCase() && admin.IsActive
      );
    } catch (error) {
      console.error("Error checking admin status:", error);
      return false;
    }
  }

  public async render(): Promise<void> {
    this.domElement.innerHTML = `
      <div class="${styles.leaveTrackerContainer}">
        <!-- Sidebar -->
        <aside class="${styles.sidebar}">
          <h2 class="${styles.title}">Leave Tracker</h2>
          <button class="${styles.menuButton} ${styles.menuActive}" data-view="dashboard">Dashboard</button>
          <button class="${styles.menuButton}" data-view="request">Request Leave</button>
          <button class="${styles.menuButton}" data-view="holidays">Gov Holidays</button>
          <button class="${styles.menuButton}" data-view="history">${this.isAdmin ? 'All Leaves' : 'My Leave History'}</button>
          ${this.isAdmin ? `<div class="${styles.adminBadge}">Admin Access</div>` : ""}
        </aside>

        <!-- Main Content -->
        <main class="${styles.mainContent}" id="mainContent">
          ${this.renderDashboard()}
        </main>
      </div>
    `;

    this.attachEventListeners();
  }

  private attachEventListeners(): void {
    // Menu buttons
    const buttons = this.domElement.querySelectorAll(`.${styles.menuButton}`);
    buttons.forEach((button: Element) => {
      button.addEventListener('click', this.handleMenuClick.bind(this));
    });

    // Leave type select
    const leaveTypeSelect = this.domElement.querySelector('#leaveType') as HTMLSelectElement;
    if (leaveTypeSelect) {
      leaveTypeSelect.addEventListener('change', this.handleLeaveTypeChange.bind(this));
    }

    // Submit button
    const submitBtn = this.domElement.querySelector('#btnSubmitLeave');
    if (submitBtn) {
      submitBtn.addEventListener('click', this.handleLeaveSubmit.bind(this));
    }
  }

  private handleMenuClick(e: Event): void {
    const target = e.currentTarget as HTMLElement;
    const view = target.getAttribute('data-view');
    if (view) {
      this.switchView(view);
    }
  }

  private handleLeaveTypeChange(e: Event): void {
    const select = e.target as HTMLSelectElement;
    const otherContainer = this.domElement.querySelector('#otherLeaveContainer') as HTMLElement;
    
    if (otherContainer) {
      otherContainer.style.display = select.value === 'Other' ? 'block' : 'none';
    }
  }

  private async handleLeaveSubmit(): Promise<void> {
    const leaveTypeSelect = this.domElement.querySelector('#leaveType') as HTMLSelectElement;
    const startDateInput = this.domElement.querySelector('#startDate') as HTMLInputElement;
    const endDateInput = this.domElement.querySelector('#endDate') as HTMLInputElement;
    const reasonInput = this.domElement.querySelector('#reason') as HTMLTextAreaElement;
    const otherLeaveInput = this.domElement.querySelector('#otherLeaveInput') as HTMLInputElement;

    if (!leaveTypeSelect || !startDateInput || !endDateInput || !reasonInput) {
      alert('Please fill all required fields');
      return;
    }

    const leaveType = leaveTypeSelect.value === 'Other' ? otherLeaveInput?.value : leaveTypeSelect.value;
    const startDate = startDateInput.value;
    const endDate = endDateInput.value;
    const reason = reasonInput.value;

    if (!leaveType || !startDate || !endDate || !reason) {
      alert('Please fill all required fields');
      return;
    }

    // Calculate number of days
    const start = new Date(startDate);
    const end = new Date(endDate);
    const numberOfDays = Math.ceil((end.getTime() - start.getTime()) / (1000 * 60 * 60 * 24)) + 1;

    // Submit to SharePoint
    try {
      const listName = "LeaveTracker team members list";
      const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
      
      const body = JSON.stringify({
        Title: `Leave Request - ${new Date().toISOString()}`,
        EmployeeEmail: this.context.pageContext.user.email,
        LeaveType: leaveType,
        StartDate: startDate,
        EndDate: endDate,
        NumberOfDays: numberOfDays,
        Reason: reason,
        RequestDate: new Date().toISOString(),
        Status: 'Pending'
      });

      await this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': ''
        },
        body: body
      });

      alert('Leave request submitted successfully!');
      
      // Reload data and switch to history view
      await this.loadTeamMembersList();
      this.switchView('history');
    } catch (error) {
      console.error('Error submitting leave request:', error);
      alert('Error submitting leave request. Please try again.');
    }
  }

  private switchView(view: string): void {
    // Update active menu button
    const buttons = this.domElement.querySelectorAll(`.${styles.menuButton}`);
    buttons.forEach((button: Element) => {
      const btnElement = button as HTMLElement;
      if (btnElement.getAttribute('data-view') === view) {
        btnElement.classList.add(styles.menuActive);
      } else {
        btnElement.classList.remove(styles.menuActive);
      }
    });

    // Update main content using cached data
    const mainContent = this.domElement.querySelector('#mainContent');
    if (mainContent) {
      switch (view) {
        case 'dashboard':
          mainContent.innerHTML = this.renderDashboard();
          break;
        case 'request':
          mainContent.innerHTML = this.renderRequestLeave(this.leaveTypes);
          this.attachEventListeners();
          break;
        case 'holidays':
          mainContent.innerHTML = this.renderGovHolidays(this.cachedHolidays);
          break;
        case 'history':
          mainContent.innerHTML = this.renderLeaveHistory(this.cachedTeamMembers);
          break;
      }
    }
  }

  private renderDashboard(): string {
    const currentUserEmail = this.context.pageContext.user.email.toLowerCase();

    // Filter only current user's leave data
    const userLeaves = this.cachedTeamMembers.filter(m =>
      m.EmployeeEmail && m.EmployeeEmail.toLowerCase() === currentUserEmail
    );

    const now = new Date();
    const oneMonthAgo = new Date();
    oneMonthAgo.setMonth(now.getMonth() - 1);

    // Only last 30 days of THIS user's leaves
    const lastMonthLeaves = userLeaves.filter(item => {
      const start = new Date(item.StartDate);
      return start >= oneMonthAgo && start <= now;
    });

    // Summary counts
    const total = lastMonthLeaves.length;
    const approved = lastMonthLeaves.filter(l => l.Status === "Approved").length;
    const rejected = lastMonthLeaves.filter(l => l.Status === "Rejected").length;
    const pending = lastMonthLeaves.filter(l => l.Status === "Pending").length;
    const totalDays = lastMonthLeaves.reduce((sum, l) => sum + (l.NumberOfDays || 0), 0);

    // Status styles
    const statusClassMap: Record<string, string> = {
      approved: styles.approved,
      pending: styles.pending,
      rejected: styles.rejected
    };

    // Recent entries
    const recentList = lastMonthLeaves
      .slice(0, 5)
      .map(l => `
        <div class="${styles.recentItem}">
          <div><strong>${escape(l.EmployeeName?.Title || 'N/A')}</strong> • ${escape(l.LeaveType)}</div>
          <div>${new Date(l.StartDate).toLocaleDateString()} → ${new Date(l.EndDate).toLocaleDateString()}</div>
          <span class="${styles.status} ${statusClassMap[l.Status?.toLowerCase()] || ""}">
            ${escape(l.Status)}
          </span>
        </div>
      `)
      .join("");

    // Final HTML
    return `
      <div class="${styles.dashboardWrapper}">
        
        <!-- Title Card -->
        <div class="${styles.card}">
          <h1>Dashboard</h1>
          <p>Summary of your leave balance and activities (Last 30 Days).</p>
        </div>

        <!-- Summary Grid -->
        <div class="${styles.summaryGrid}">
          <div class="${styles.summaryCard}">
            <div>
              <h3>${total}</h3>
              <p>Total Requests</p>
            </div>
          </div>

          <div class="${styles.summaryCard}">
            <div>
              <h3>${approved}</h3>
              <p>Approved</p>
            </div>
          </div>

          <div class="${styles.summaryCard}">
            <div>
              <h3>${pending}</h3>
              <p>Pending</p>
            </div>
          </div>

          <div class="${styles.summaryCard}">
            <div>
              <h3>${rejected}</h3>
              <p>Rejected</p>
            </div>
          </div>

          <div class="${styles.summaryCard}">
            <div>
              <h3>${totalDays}</h3>
              <p>Total Days</p>
            </div>
          </div>
        </div>

        <!-- Recent Leaves -->
        <div class="${styles.card}">
          <h2>Recent Leave Entries</h2>
          <div class="${styles.recentList}">
            ${recentList || "<p>No leave data for the last 30 days.</p>"}
          </div>
        </div>

      </div>
    `;
  }

  private renderRequestLeave(leaveTypes: string[]): string {
    const updatedTypes = Array.from(new Set([...leaveTypes, "Other"]));

    const optionsHtml = updatedTypes
      .map(type => `<option value="${escape(type)}">${escape(type)}</option>`)
      .join('');

    return `
      <div class="${styles.card}">
        <h1>Request Leave</h1>

        <div class="${styles.formGroup}">
          <label class="${styles.label}">Leave Type</label>
          <select class="${styles.select}" id="leaveType">
            ${optionsHtml}
          </select>
        </div>

        <!-- Hidden input for Other -->
        <div class="${styles.formGroup}" id="otherLeaveContainer" style="display:none;">
          <label class="${styles.label}">Enter Leave Type</label>
          <input type="text" class="${styles.input}" id="otherLeaveInput" placeholder="Enter custom leave type" />
        </div>

        <div class="${styles.formGroup}">
          <label class="${styles.label}">Start Date</label>
          <input type="date" class="${styles.input}" id="startDate" />
        </div>

        <div class="${styles.formGroup}">
          <label class="${styles.label}">End Date</label>
          <input type="date" class="${styles.input}" id="endDate" />
        </div>

        <div class="${styles.formGroup}">
          <label class="${styles.label}">Reason</label>
          <textarea class="${styles.textarea}" rows="3" id="reason" placeholder="Enter reason for leave"></textarea>
        </div>

        <button class="${styles.submitBtn}" id="btnSubmitLeave">
          Submit Request
        </button>
      </div>
    `;
  }

  private renderGovHolidays(holidays: IGovernmentHoliday[]): string {
    // Filter active holidays and format them
    const holidayItems = holidays
      .filter(h => h.IsActive)
      .map(h => {
        const date = new Date(h.HolidayDate);
        const day = date.getDate();
        const month = date.toLocaleDateString('en-US', { month: 'short' });

        return `
          <div class="${styles.holidayItemBox}">
            <div class="${styles.holidayDate}">${day} ${month}</div>
            <div class="${styles.holidayName}">${escape(h.Title)}</div>
          </div>
        `;
      }).join('');

    // Only show Add button to admins
    const adminControls = this.isAdmin ? `
      <div class="${styles.adminControls}">
        <button class="${styles.submitBtn}" id="btnAddHoliday">Add Holiday</button>
      </div>
    ` : '';

    return `
      <div class="${styles.card}">
        <h1>Government Holidays</h1>
        <div class="${styles.holidayList}">
          ${holidayItems || '<p>No holidays available.</p>'}
        </div>
        ${adminControls}
      </div>
    `;
  }

  private renderLeaveHistory(teamMembers: ITeamMember[]): string {
    const currentUserEmail = this.context.pageContext.user.email.toLowerCase();

    // Filter leaves for current user only
    const userLeaves = teamMembers.filter(member =>
      member.EmployeeEmail && member.EmployeeEmail.toLowerCase() === currentUserEmail
    );

    // Helper function to get status class
    const getStatusClass = (status: string): string => {
      const statusLower = status?.toLowerCase();
      switch (statusLower) {
        case 'approved':
          return styles.approved;
        case 'pending':
          return styles.pending;
        case 'rejected':
          return styles.rejected;
        default:
          return '';
      }
    };

    const rows = userLeaves.length > 0 ? userLeaves.map(leave => `
      <tr>
        <td>${escape(leave.LeaveType)}</td>
        <td>${escape(leave.StartDate?.split('T')[0] || 'N/A')}</td>
        <td>${escape(leave.EndDate?.split('T')[0] || 'N/A')}</td>
        <td><span class="${styles.statusBadge} ${getStatusClass(leave.Status)}">${escape(leave.Status)}</span></td>
      </tr>
    `).join('') : '<tr><td colspan="4" style="text-align: center;">No leave history available</td></tr>';

    return `
      <div class="${styles.card}">
        <h1>Leave History</h1>
        <table class="${styles.table}">
          <thead>
            <tr>
              <th>Type</th>
              <th>From</th>
              <th>To</th>
              <th>Status</th>
            </tr>
          </thead>
          <tbody>
            ${rows}
          </tbody>
        </table>
      </div>
    `;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Leave Tracker Settings'
          },
          groups: [
            {
              groupName: 'Basic Settings',
              groupFields: [
                PropertyPaneTextField('description', {
                  label: 'Description'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}