import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient } from '@microsoft/sp-http';

import styles from './LeaveTrackerWebPart.module.scss';
import {
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';

export interface ILeaveTrackerWebPartProps {
  description: string;
  adminListName: string;
  teamMembersListName: string;
  holidaysListName: string;
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
interface ISiteList {
  Title: string;
  Id: string;
}
export default class LeaveTrackerWebPart extends BaseClientSideWebPart<ILeaveTrackerWebPartProps> {
  private leaveTypes: string[] = [];
  private cachedAdmins: IAdmin[] = [];
  private cachedTeamMembers: ITeamMember[] = [];
  private cachedHolidays: IGovernmentHoliday[] = [];
  private isAdmin: boolean = false;
  private holidayTypes: string[] = [];

  private siteListsCache: ISiteList[] = [];
  private listDropdownOptions: IPropertyPaneDropdownOption[] = [];

  // 4. ADD THIS METHOD (after closeSidePanel method)
  private async getSiteListNames(): Promise<void> {
    try {
      const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$select=Title,Id&$filter=Hidden eq false&$orderby=Title`;
      const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      const data = await response.json();

      if (data.value && data.value.length > 0) {
        this.siteListsCache = data.value;
        console.log("Site Lists:", this.siteListsCache);

        // Convert to PropertyPane dropdown options
        this.listDropdownOptions = this.siteListsCache.map((list: ISiteList) => ({
          key: list.Title,
          text: list.Title
        }));

        console.log("Dropdown Options:", this.listDropdownOptions);
      }
    } catch (error) {
      console.error("Error fetching site lists:", error);
      this.listDropdownOptions = [];
    }
  }

  // 5. UPDATE onInit() METHOD (add getSiteListNames() to Promise.all)
  protected async onInit(): Promise<void> {
    await Promise.all([
      this.loadAdminList(),
      this.loadTeamMembersList(),
      this.loadHolidaysList(),
      this.loadLeaveTypeChoices(),
      this.loadHolidayTypeChoices(),
      this.getSiteListNames()  // ADD THIS LINE
    ]);

    this.isAdmin = await this.checkUserAdmin();
    return super.onInit();
  }

  private getCurrentUserEmail(): string {
    // return this.context.pageContext.user.email?.toLowerCase() || "";
    return "jacob.yeldhos@aciesinnovations.com";
  }

  private async loadHolidayTypeChoices(): Promise<void> {
    const listName = this.properties.holidaysListName || "Government Holidays";
    try {
      const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/fields?$filter=InternalName eq 'HolidayType'`;
      const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      const data = await response.json();

      if (data.value && data.value.length > 0) {
        this.holidayTypes = data.value[0].Choices || [];
      }
    } catch (error) {
      console.error("Error fetching HolidayType choices:", error);
      this.holidayTypes = [];
    }
  }

  private async loadAdminList(): Promise<void> {
    const listName = this.properties.adminListName || "LeaveTracker Admin List";
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
    const listName = this.properties.teamMembersListName || "LeaveTracker team members list";
    try {
      const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=*,EmployeeName/Title,EmployeeName/EMail&$expand=EmployeeName&$top=5000`;
      const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      const data = await response.json();
      this.cachedTeamMembers = data.value;
      console.log("Team Members List Data (Total):", this.cachedTeamMembers.length);
      console.log("Team Members Data:", this.cachedTeamMembers);
    } catch (error) {
      console.error("Error loading Team Members List:", error);
      this.cachedTeamMembers = [];
    }
  }

  private async loadHolidaysList(): Promise<void> {
    const listName = this.properties.holidaysListName || "Government Holidays";
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
    const listName = this.properties.teamMembersListName || "LeaveTracker team members list";
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
    const userEmail = this.getCurrentUserEmail();

    console.log("Checking admin status for:", userEmail);
    console.log("Cached admins:", this.cachedAdmins);

    try {
      const isAdmin = this.cachedAdmins.some((admin: IAdmin) =>
        admin.AdminEmail && admin.AdminEmail.toLowerCase() === userEmail && admin.IsActive
      );

      console.log("Is Admin Result:", isAdmin);
      return isAdmin;
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

  private handleMenuClick = (e: Event): void => {
    const target = e.currentTarget as HTMLElement;
    const view = target.getAttribute('data-view');
    if (view) {
      this.switchView(view);
    }
  }

  private handleLeaveTypeChange = (e: Event): void => {
    const select = e.target as HTMLSelectElement;
    const otherContainer = this.domElement.querySelector('#otherLeaveContainer') as HTMLElement;

    if (otherContainer) {
      otherContainer.style.display = select.value === 'Other' ? 'block' : 'none';
    }
  }

  private async getUserIdByEmail(email: string): Promise<number | null> {
    try {
      const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/siteusers?$filter=Email eq '${email}'`;
      const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      const data = await response.json();

      if (data.value && data.value.length > 0) {
        console.log("User ID found:", data.value[0].Id);
        return data.value[0].Id;
      }

      console.log("User ID not found for email:", email);
      return null;
    } catch (error) {
      console.error("Error fetching user ID:", error);
      return null;
    }
  }

  private extractNameFromEmail(email: string): string {
    const namePart = email.split('@')[0];
    const parts = namePart.split(/[._]/);

    const formattedName = parts
      .map(part => part.charAt(0).toUpperCase() + part.slice(1).toLowerCase())
      .join(' ');

    return formattedName;
  }

  private handleLeaveSubmit = async (): Promise<void> => {
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

    const start = new Date(startDate);
    const end = new Date(endDate);
    const numberOfDays = Math.ceil((end.getTime() - start.getTime()) / (1000 * 60 * 60 * 24)) + 1;

    try {
      const listName = this.properties.teamMembersListName || "LeaveTracker team members list";
      const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`;

      const currentUserEmail = this.getCurrentUserEmail();
      const userId = await this.getUserIdByEmail(currentUserEmail);

      const bodyData: any = {
        Title: `Leave Request - ${new Date().toISOString()}`,
        EmployeeEmail: currentUserEmail,
        LeaveType: leaveType,
        StartDate: startDate,
        EndDate: endDate,
        NumberOfDays: numberOfDays,
        Reason: reason,
        RequestDate: new Date().toISOString(),
        Status: 'Pending'
      };

      if (userId) {
        bodyData.EmployeeNameId = userId;
        console.log("Setting EmployeeName with User ID:", userId);
      } else {
        const extractedName = this.extractNameFromEmail(currentUserEmail);
        console.log("User ID not found, extracted name from email:", extractedName);
      }

      const body = JSON.stringify(bodyData);

      await this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': ''
        },
        body: body
      });

      console.log("Leave request submitted successfully", bodyData);
      await this.loadTeamMembersList();
      this.switchView('history');
    } catch (error) {
      console.error('Error submitting leave request:', error);
      alert('Error submitting leave request. Please try again.');
    }
  }


  private switchView(view: string): void {

    const buttons = this.domElement.querySelectorAll(`.${styles.menuButton}`);
    buttons.forEach((button: Element) => {
      const btnElement = button as HTMLElement;
      if (btnElement.getAttribute('data-view') === view) {
        btnElement.classList.add(styles.menuActive);
      } else {
        btnElement.classList.remove(styles.menuActive);
      }
    });

    const mainContent = this.domElement.querySelector('#mainContent');
    if (mainContent) {
      switch (view) {
        case 'dashboard':
          mainContent.innerHTML = this.renderDashboard();
          break;
        case 'request':
          mainContent.innerHTML = this.renderRequestLeave(this.leaveTypes);
          break;
        case 'holidays':
          mainContent.innerHTML = this.renderGovHolidays(this.cachedHolidays);
          break;
        case 'history':
          mainContent.innerHTML = this.renderLeaveHistory(this.cachedTeamMembers, 'all', 'all');
          break;
      }
      this.attachEventListeners();
    }
  }

  private renderDashboard(): string {
    const currentUserEmail = this.getCurrentUserEmail();

    const userLeaves = this.cachedTeamMembers.filter(m =>
      m.EmployeeEmail && m.EmployeeEmail.toLowerCase() === currentUserEmail
    );

    const now = new Date();
    const oneMonthAgo = new Date();
    oneMonthAgo.setMonth(now.getMonth() - 1);

    const lastMonthLeaves = userLeaves.filter(item => {
      const start = new Date(item.StartDate);
      return start >= oneMonthAgo && start <= now;
    });

    const total = lastMonthLeaves.length;
    const approved = lastMonthLeaves.filter(l => l.Status === "Approve").length;
    const rejected = lastMonthLeaves.filter(l => l.Status === "Rejected").length;
    const pending = lastMonthLeaves.filter(l => l.Status === "Pending").length;

    const statusClassMap: Record<string, string> = {
      approved: styles.approved,
      pending: styles.pending,
      rejected: styles.rejected
    };

    const recentList = lastMonthLeaves
      .slice(0, 5)
      .map(l => `
        <div class="${styles.recentItem}">
          <div>${escape(l.LeaveType)}</div>
          <div>${new Date(l.StartDate).toLocaleDateString()} → ${new Date(l.EndDate).toLocaleDateString()}</div>
          <span class="${styles.status} ${statusClassMap[l.Status?.toLowerCase()] || ""}">
            ${escape(l.Status)}
          </span>
        </div>
      `)
      .join("");

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
              <h5>Flexible </h5>
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

    const holidayTypeOptions = this.holidayTypes
      .map(type => `<option value="${escape(type)}">${escape(type)}</option>`)
      .join('');

    const adminControls = this.isAdmin ? `
    <div class="${styles.adminControls}">
      <button class="${styles.submitBtn}" id="btnAddHoliday">+ Add Holiday</button>
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

    <!-- Side Panel for Adding Holiday -->
    <div class="${styles.sidePanel}" id="holidaySidePanel">
      <div class="${styles.sidePanelOverlay}" id="sidePanelOverlay"></div>
      <div class="${styles.sidePanelContent}">
        <div class="${styles.sidePanelHeader}">
          <h2>Add New Holiday</h2>
          <button class="${styles.closePanelBtn}" id="closePanelBtn">&times;</button>
        </div>

        <div class="${styles.sidePanelBody}">
          <div class="${styles.formGroup}">
            <label class="${styles.label}">Holiday Name *</label>
            <input 
              type="text" 
              class="${styles.input}" 
              id="holidayTitle" 
              placeholder="e.g., Independence Day"
              required
            />
          </div>

          <div class="${styles.formGroup}">
            <label class="${styles.label}">Holiday Date *</label>
            <input 
              type="date" 
              class="${styles.input}" 
              id="holidayDate"
              required
            />
          </div>

          <div class="${styles.formGroup}">
            <label class="${styles.label}">Holiday Type *</label>
            <select class="${styles.select}" id="holidayType" required>
              <option value="">Select Type</option>
              ${holidayTypeOptions}
            </select>
          </div>

          <div class="${styles.formGroup}">
            <label class="${styles.label}">Description</label>
            <textarea 
              class="${styles.textarea}" 
              rows="4" 
              id="holidayDescription" 
              placeholder="Enter holiday description (optional)"
            ></textarea>
          </div>

          <div class="${styles.formGroup}">
            <label class="${styles.checkboxLabel}">
              <input 
                type="checkbox" 
                id="holidayIsActive" 
                checked
                class="${styles.checkbox}"
              />
              <span>Active Holiday</span>
            </label>
          </div>
        </div>

        <div class="${styles.sidePanelFooter}">
          <button class="${styles.cancelBtn}" id="cancelHolidayBtn">Cancel</button>
          <button class="${styles.submitBtn}" id="saveHolidayBtn">Save Holiday</button>
        </div>
      </div>
    </div>
  `;
  }

  private attachEventListeners(): void {
    // Menu buttons
    const buttons = this.domElement.querySelectorAll(`.${styles.menuButton}`);
    buttons.forEach((button: Element) => {
      button.addEventListener('click', this.handleMenuClick);
    });

    // Leave type select
    const leaveTypeSelect = this.domElement.querySelector('#leaveType') as HTMLSelectElement;
    if (leaveTypeSelect) {
      leaveTypeSelect.addEventListener('change', this.handleLeaveTypeChange);
    }

    // Submit leave button
    const submitBtn = this.domElement.querySelector('#btnSubmitLeave');
    if (submitBtn) {
      submitBtn.addEventListener('click', this.handleLeaveSubmit);
    }

    // Time filter select
    const timeFilter = this.domElement.querySelector('#timeFilter') as HTMLSelectElement;
    if (timeFilter) {
      timeFilter.addEventListener('change', (e: Event) => {
        const select = e.target as HTMLSelectElement;

        const viewModeButtons = this.domElement.querySelectorAll('[data-view-mode]');
        let currentViewMode = 'all';
        viewModeButtons.forEach(btn => {
          if (btn.classList.contains(styles.toggleBtnActive)) {
            currentViewMode = btn.getAttribute('data-view-mode') || 'all';
          }
        });

        const mainContent = this.domElement.querySelector('#mainContent');
        if (mainContent) {
          mainContent.innerHTML = this.renderLeaveHistory(this.cachedTeamMembers, select.value, currentViewMode);
          this.attachEventListeners();
        }
      });
    }

    // View mode toggle buttons (All Requests / My Requests)
    const viewModeButtons = this.domElement.querySelectorAll('[data-view-mode]');
    viewModeButtons.forEach((button: Element) => {
      button.addEventListener('click', (e: Event) => {
        const btn = e.currentTarget as HTMLElement;
        const viewMode = btn.getAttribute('data-view-mode') || 'all';

        console.log("View mode clicked:", viewMode);

        const timeFilterSelect = this.domElement.querySelector('#timeFilter') as HTMLSelectElement;
        const currentFilter = timeFilterSelect ? timeFilterSelect.value : 'all';

        const mainContent = this.domElement.querySelector('#mainContent');
        if (mainContent) {
          mainContent.innerHTML = this.renderLeaveHistory(this.cachedTeamMembers, currentFilter, viewMode);
          this.attachEventListeners();
        }
      });
    });

    // Month tabs with proper filtering
    const monthTabs = this.domElement.querySelectorAll('[data-month]');
    monthTabs.forEach((tab: Element) => {
      tab.addEventListener('click', (e: Event) => {
        const button = e.currentTarget as HTMLElement;
        const month = parseInt(button.getAttribute('data-month') || '0');

        console.log("Month clicked:", month);

        monthTabs.forEach(t => t.classList.remove(styles.monthTabActive));
        button.classList.add(styles.monthTabActive);

        const rows = this.domElement.querySelectorAll(`.${styles.tableRow}`);
        rows.forEach((row: Element) => {
          const rowElement = row as HTMLElement;
          const rowMonth = parseInt(rowElement.getAttribute('data-month') || '-1');

          if (rowMonth === month) {
            rowElement.style.display = '';
          } else {
            rowElement.style.display = 'none';
          }
        });
      });
    });

    // Search input for admin view
    const searchInput = this.domElement.querySelector('#searchEmployee') as HTMLInputElement;
    if (searchInput) {
      searchInput.addEventListener('input', (e: Event) => {
        const input = e.target as HTMLInputElement;
        const searchTerm = input.value.toLowerCase();
        const rows = this.domElement.querySelectorAll(`.${styles.tableRow}`);

        rows.forEach((row: Element) => {
          const rowElement = row as HTMLElement;
          const employeeName = rowElement.getAttribute('data-employee')?.toLowerCase() || '';
          const rowText = rowElement.textContent?.toLowerCase() || '';

          if (employeeName.includes(searchTerm) || rowText.includes(searchTerm)) {
            rowElement.style.display = '';
          } else {
            rowElement.style.display = 'none';
          }
        });
      });
    }

    // Add Holiday button
    const addHolidayBtn = this.domElement.querySelector('#btnAddHoliday');
    if (addHolidayBtn) {
      addHolidayBtn.addEventListener('click', () => {
        this.openSidePanel();
      });
    }

    // Close panel button (X button)
    const closePanelBtn = this.domElement.querySelector('#closePanelBtn');
    if (closePanelBtn) {
      closePanelBtn.addEventListener('click', () => {
        this.closeSidePanel();
      });
    }

    // Cancel button in side panel
    const cancelHolidayBtn = this.domElement.querySelector('#cancelHolidayBtn');
    if (cancelHolidayBtn) {
      cancelHolidayBtn.addEventListener('click', () => {
        this.closeSidePanel();
      });
    }

    // Save holiday button
    const saveHolidayBtn = this.domElement.querySelector('#saveHolidayBtn');
    if (saveHolidayBtn) {
      saveHolidayBtn.addEventListener('click', () => {
        this.handleHolidaySubmit();
      });
    }

    // Close panel when clicking overlay (backdrop)
    const sidePanelOverlay = this.domElement.querySelector('#sidePanelOverlay');
    if (sidePanelOverlay) {
      sidePanelOverlay.addEventListener('click', () => {
        this.closeSidePanel();
      });
    }
  }

  private async handleHolidaySubmit(): Promise<void> {
    const titleInput = this.domElement.querySelector('#holidayTitle') as HTMLInputElement;
    const dateInput = this.domElement.querySelector('#holidayDate') as HTMLInputElement;
    const typeSelect = this.domElement.querySelector('#holidayType') as HTMLSelectElement;
    const descriptionInput = this.domElement.querySelector('#holidayDescription') as HTMLTextAreaElement;
    const isActiveCheckbox = this.domElement.querySelector('#holidayIsActive') as HTMLInputElement;

    if (!titleInput || !dateInput || !typeSelect) {
      alert('Required fields are missing');
      return;
    }

    const title = titleInput.value.trim();
    const holidayDate = dateInput.value;
    const holidayType = typeSelect.value;
    const description = descriptionInput?.value.trim() || '';
    const isActive = isActiveCheckbox?.checked ?? true;

    if (!title || !holidayDate || !holidayType) {
      alert('Please fill all required fields (Holiday Name, Date, and Type)');
      return;
    }

    try {
      const listName = this.properties.holidaysListName || "Government Holidays";
      const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`;

      const body = JSON.stringify({
        Title: title,
        HolidayDate: holidayDate,
        HolidayType: holidayType,
        Description: description,
        IsActive: isActive
      });

      const response = await this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': ''
        },
        body: body
      });

      if (response.ok) {
        console.log("Holiday added successfully");
        alert('Holiday added successfully!');

        this.closeSidePanel();

        await this.loadHolidaysList();
        this.switchView('holidays');
      } else {
        throw new Error('Failed to add holiday');
      }
    } catch (error) {
      console.error('Error adding holiday:', error);
      alert('Error adding holiday. Please try again.');
    }
  }

  private openSidePanel(): void {
    const sidePanel = this.domElement.querySelector('#holidaySidePanel') as HTMLElement;
    if (sidePanel) {
      sidePanel.classList.add(styles.sidePanelOpen);
      document.body.style.overflow = 'hidden';
    }
  }

  private closeSidePanel(): void {
    const sidePanel = this.domElement.querySelector('#holidaySidePanel') as HTMLElement;
    if (sidePanel) {
      sidePanel.classList.remove(styles.sidePanelOpen);
      document.body.style.overflow = '';

      const titleInput = this.domElement.querySelector('#holidayTitle') as HTMLInputElement;
      const dateInput = this.domElement.querySelector('#holidayDate') as HTMLInputElement;
      const typeSelect = this.domElement.querySelector('#holidayType') as HTMLSelectElement;
      const descriptionInput = this.domElement.querySelector('#holidayDescription') as HTMLTextAreaElement;
      const isActiveCheckbox = this.domElement.querySelector('#holidayIsActive') as HTMLInputElement;

      if (titleInput) titleInput.value = '';
      if (dateInput) dateInput.value = '';
      if (typeSelect) typeSelect.value = '';
      if (descriptionInput) descriptionInput.value = '';
      if (isActiveCheckbox) isActiveCheckbox.checked = true;
    }
  }

  private renderLeaveHistory(teamMembers: ITeamMember[], filter: string = 'all', viewMode: string = 'all'): string {
    const currentUserEmail = this.getCurrentUserEmail();

    console.log("Rendering leave history - Is Admin:", this.isAdmin);
    console.log("Total team members:", teamMembers.length);
    console.log("View Mode:", viewMode);

    let filteredLeaves: ITeamMember[];

    if (viewMode === 'mine') {
      filteredLeaves = teamMembers.filter(member =>
        member.EmployeeEmail && member.EmployeeEmail.toLowerCase() === currentUserEmail
      );
      console.log("My Requests view:", filteredLeaves.length);
    } else if (this.isAdmin) {
      filteredLeaves = [...teamMembers];
      console.log("Admin view - showing all leaves:", filteredLeaves.length);
    } else {
      filteredLeaves = teamMembers.filter(member =>
        member.EmployeeEmail && member.EmployeeEmail.toLowerCase() === currentUserEmail
      );
      console.log("User view - showing own leaves:", filteredLeaves.length);
    }

    const filterDate = this.getFilterDate(filter);
    filteredLeaves = filteredLeaves.filter(leave => {
      const requestDate = new Date(leave.RequestDate || leave.StartDate);
      return requestDate >= filterDate;
    });

    console.log("After filter:", filteredLeaves.length);

    filteredLeaves.sort((a, b) => {
      const dateA = new Date(a.RequestDate || a.StartDate).getTime();
      const dateB = new Date(b.RequestDate || b.StartDate).getTime();
      return dateB - dateA;
    });

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const isOnLeaveToday = (leave: ITeamMember): boolean => {
      const startDate = new Date(leave.StartDate);
      const endDate = new Date(leave.EndDate);
      startDate.setHours(0, 0, 0, 0);
      endDate.setHours(0, 0, 0, 0);

      return leave.Status === 'Approve' && startDate <= today && endDate >= today;
    };

    const getStatusClass = (status: string): string => {
      const statusLower = status?.toLowerCase();
      switch (statusLower) {
        case 'approve':
        case 'approved':
          return styles.statusActive;
        case 'pending':
          return styles.statusPending;
        case 'reject':
        case 'rejected':
          return styles.statusClosed;
        default:
          return styles.statusOffline;
      }
    };

    const currentYear = new Date().getFullYear();
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const currentMonth = new Date().getMonth();

    const rows = filteredLeaves.length > 0 ? filteredLeaves.map(leave => {
      const startDate = new Date(leave.StartDate);
      const endDate = new Date(leave.EndDate);
      const requestDate = new Date(leave.RequestDate || leave.StartDate);

      const formattedRequestDate = requestDate.toLocaleDateString('en-US', {
        year: 'numeric',
        month: 'short',
        day: 'numeric'
      });

      const formattedStartDate = startDate.toLocaleDateString('en-US', {
        month: 'short',
        day: 'numeric'
      });

      const formattedEndDate = endDate.toLocaleDateString('en-US', {
        month: 'short',
        day: 'numeric',
        year: 'numeric'
      });

      const dateRange = `${formattedStartDate} - ${formattedEndDate}`;
      const numberOfDays = leave.NumberOfDays || 0;

      let approverInfo = '-';
      if (leave.Status === 'Approve' && leave.ApproveDate) {
        approverInfo = 'Approved';
      } else if (leave.Status === 'Rejected' && leave.RejectionReason) {
        approverInfo = `Rejected: ${leave.RejectionReason}`;
      }

      const onLeaveNow = isOnLeaveToday(leave);
      const rowClass = onLeaveNow ? `${styles.tableRow} ${styles.onLeaveRow}` : styles.tableRow;

      return `
      <tr class="${rowClass}" data-month="${startDate.getMonth()}" data-employee="${escape(leave.EmployeeName?.Title || leave.EmployeeEmail || '')}">
        <td class="${styles.tableCell}">
          <div class="${styles.employeeInfo}">
            <div class="${styles.avatar} ${onLeaveNow ? styles.avatarOnLeave : ''}">${(leave.EmployeeName?.Title || leave.EmployeeEmail || 'U')[0].toUpperCase()}</div>
            <span class="${styles.employeeName}">
              ${escape(leave.EmployeeName?.Title || leave.EmployeeEmail || 'N/A')}
              ${onLeaveNow ? '<span class="' + styles.onLeaveBadge + '">🟢 On Leave</span>' : ''}
            </span>
          </div>
        </td>
        <td class="${styles.tableCell}">${escape(leave.LeaveType)}</td>
        <td class="${styles.tableCell}">
          <span class="${getStatusClass(leave.Status)}">${escape(leave.Status)}</span>
        </td>
        <td class="${styles.tableCell}">${dateRange} (${numberOfDays}d)</td>
        <td class="${styles.tableCell}" style="max-width: 200px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${escape(leave.Reason || '-')}">${escape(leave.Reason || '-')}</td>
        <td class="${styles.tableCell}">${formattedRequestDate}</td>
        <td class="${styles.tableCell}">${approverInfo}</td>
      </tr>
    `;
    }).join('') : `<tr><td colspan="7" class="${styles.emptyState}">No leave records found</td></tr>`;

    const filterOptions = [
      { value: 'all', label: 'All time' },
      { value: 'month', label: 'This month' },
      { value: 'week', label: 'This week' },
      { value: 'today', label: 'Today' }
    ];

    return `
    <div class="${styles.leaveHistoryContainer}">
      <!-- Header with filters and view toggle -->
      <div class="${styles.tableHeader}">
        <div class="${styles.filterSection}">
          ${this.isAdmin ? `
            <div class="${styles.viewToggle}">
              <button 
                class="${styles.toggleBtn} ${viewMode === 'all' ? styles.toggleBtnActive : ''}" 
                data-view-mode="all"
              >
                All Requests
              </button>
              <button 
                class="${styles.toggleBtn} ${viewMode === 'mine' ? styles.toggleBtnActive : ''}" 
                data-view-mode="mine"
              >
                My Requests
              </button>
            </div>
          ` : ''}
          
          <div style="display: flex; gap: 10px; align-items: center;">
            <span class="${styles.filterLabel}">Filter by</span>
            <select class="${styles.filterSelect}" id="timeFilter">
              ${filterOptions.map(opt => `
                <option value="${opt.value}" ${filter === opt.value ? 'selected' : ''}>${opt.label}</option>
              `).join('')}
            </select>
            
            ${this.isAdmin && viewMode === 'all' ? `
              <input 
                type="text" 
                id="searchEmployee" 
                class="${styles.searchInput}" 
                placeholder="🔍 Search employee..."
              />
            ` : ''}
          </div>
        </div>
      </div>

      <!-- Month tabs -->
      <div class="${styles.monthTabs}">
        <div class="${styles.yearLabel}">${currentYear}</div>
        ${months.map((month, index) => `
          <button 
            class="${styles.monthTab} ${index === currentMonth ? styles.monthTabActive : ''}" 
            data-month="${index}"
          >
            ${month}
          </button>
        `).join('')}
      </div>

      <!-- Table -->
      <div class="${styles.modernTable}">
        <table style="width: 100%;">
          <thead>
            <tr class="${styles.tableHeaderRow}">
              <th class="${styles.tableHeader}">NAME</th>
              <th class="${styles.tableHeader}">LEAVE TYPE</th>
              <th class="${styles.tableHeader}">STATUS</th>
              <th class="${styles.tableHeader}">LEAVE PERIOD</th>
              <th class="${styles.tableHeader}">REASON</th>
              <th class="${styles.tableHeader}">REQUESTED ON</th>
              <th class="${styles.tableHeader}">APPROVAL INFO</th>
            </tr>
          </thead>
          <tbody>
            ${rows}
          </tbody>
        </table>
      </div>
    </div>
  `;
  }

  private getFilterDate(filter: string): Date {
    const now = new Date();

    switch (filter) {
      case 'today':
        return new Date(now.setHours(0, 0, 0, 0));
      case 'week':
        const weekAgo = new Date();
        weekAgo.setDate(weekAgo.getDate() - 7);
        return weekAgo;
      case 'month':
        const monthAgo = new Date();
        monthAgo.setMonth(monthAgo.getMonth() - 1);
        return monthAgo;
      case 'quarter':
        const quarterAgo = new Date();
        quarterAgo.setMonth(quarterAgo.getMonth() - 3);
        return quarterAgo;
      case 'year':
        const yearAgo = new Date();
        yearAgo.setFullYear(yearAgo.getFullYear() - 1);
        return yearAgo;
      case 'all':
      default:
        return new Date(0);
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Leave Tracker Configuration'
          },
          groups: [
            {
              groupName: 'SharePoint Lists Configuration',
              groupFields: [
                PropertyPaneDropdown('adminListName', {
                  label: 'Admin List Name',
                  options: this.listDropdownOptions,
                  selectedKey: this.properties.adminListName || 'LeaveTracker Admin List'
                }),
                PropertyPaneDropdown('teamMembersListName', {
                  label: 'Team Members List Name',
                  options: this.listDropdownOptions,
                  selectedKey: this.properties.teamMembersListName || 'LeaveTracker team members list'
                }),
                PropertyPaneDropdown('holidaysListName', {
                  label: 'Holidays List Name',
                  options: this.listDropdownOptions,
                  selectedKey: this.properties.holidaysListName || 'Government Holidays'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
