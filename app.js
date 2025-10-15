// Super Live CrowdStrike EDR Dashboard - Enhanced Interactive JavaScript
class SuperLiveDashboard {
  constructor() {
    this.excelData = null; // Store Excel data directly
    this.dashboardData = null; // Converted dashboard format
    this.osChart = null;
    this.animationSpeed = 300;
    this.updateInterval = 5000; // 5 seconds for super live feel
    this.countAnimationDuration = 2000;
    this.isLoading = false;
    this.tooltipElement = null;

    const tableBody = document.querySelector("#dataTable tbody");
    const prevBtn = document.getElementById("prevBtn");
    const nextBtn = document.getElementById("nextBtn");
    this.pageInfo = document.getElementById("pageInfo");
    const recordsSelect = document.getElementById("#recordsPerPage");
    this.totalDevicesCount = 0;

    // pagination change
    this.currentPage = 1;
    this.recordsPerPage = 10;

    // Enhanced chart colors for dark theme
    this.chartColors = {
      primary: "rgba(59, 130, 246, 0.8)",
      secondary: "rgba(99, 102, 241, 0.8)",
      success: "rgba(16, 185, 129, 0.8)",
      warning: "rgba(245, 158, 11, 0.8)",
      danger: "rgba(239, 68, 68, 0.8)",
      info: "rgba(6, 182, 212, 0.8)",
    };

    this.chartBorderColors = {
      primary: "rgba(59, 130, 246, 1)",
      secondary: "rgba(99, 102, 241, 1)",
      success: "rgba(16, 185, 129, 1)",
      warning: "rgba(245, 158, 11, 1)",
      danger: "rgba(239, 68, 68, 1)",
      info: "rgba(6, 182, 212, 1)",
    };

    this.init();
  }

  // Enhanced initialization
  init() {
    this.showLoadingOverlay();
    this.setupEventListeners();
    this.initializeTooltips();
    this.initializeLiveClock();
    this.loadDashboardData();
  }

  setPaginatorElements(currPage, recPerPage) {
    this.currentPage = parseInt(currPage);
    this.recordsPerPage = parseInt(recPerPage);
  }

  paginatorPreBtn() {
    if (this.currentPage > 1) {
      this.currentPage--;
      this.populateDevicesTable();
    }
  }

  paginatorNxtBtn() {
    if (this.currentPage * this.recordsPerPage < this.totalDevicesCount) {
      this.currentPage++;
      this.populateDevicesTable();
    }
  }

  // Load data from the Excel-based JSON file
  async loadDashboardData() {
    try {
      this.showLoadingOverlay();

      // Load the Excel data JSON
      // const response = await fetch("./dashboard_data.json");
      const response = await fetch(
        "https://api.jsonsilo.com/0d37d25c-00b5-4789-96f6-1eac7bb8e002",
        {
          headers: {
            "X-SILO-KEY": "80fGPhlcMg8qSZdmZy9uE9YKzJqKo2dKW0QvWGkf2n",
            "Content-Type": "application/json",
          },
        }
      );

      // console.log("Merry Christmas", response);

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      this.excelData = await response.json();

      // Convert Excel data to dashboard format
      this.dashboardData = this.convertExcelToDashboardFormat();

      await this.updateDashboard();
      this.hideLoadingOverlay();
      this.showNotification("Excel data loaded successfully", "success", 3000);
    } catch (error) {
      console.error("Error loading Excel data:", error);
      this.excelData = this.getFallbackData();
      this.dashboardData = this.convertExcelToDashboardFormat();
      await this.updateDashboard();
      this.hideLoadingOverlay();
      this.showNotification(
        "Using offline data - Excel file not accessible",
        "warning",
        5000
      );
    }
  }

  // Convert Excel sheets to dashboard format
  convertExcelToDashboardFormat() {
    if (!this.excelData) {
      console.log("this executes");
      return this.getEmptyDashboardFormat();
    }

    // const sheets = this.excelData.sheets;
    const sheets = this.excelData;

    // Get metrics from Dashboard_Metrics sheet
    const metricsData = sheets.Dashboard_Metrics || [];
    const metrics = {};
    metricsData.forEach((item) => {
      const key = item.Metric?.toLowerCase().replace(/\s+/g, "");
      metrics[key] = item.Value || item.Percentage || 0;
    });

    // Get OS distribution
    const osData = sheets.OS_Distribution || [];
    const osDistribution = osData.map((item) => ({
      name: item["Operating System"] || "Unknown",
      value: parseInt(item.Count) || 0,
      percentage: parseFloat(item.Percentage || 0),
      version: item["Operating System"]?.includes("11") ? "22H2" : "22H2",
      security: item["Operating System"] === "macOS" ? "High" : "High",
    }));

    // Get status breakdown
    const statusData = sheets.Status_Breakdown || [];
    const statusBreakdown = {};
    statusData.forEach((item) => {
      const key =
        item["Status Category"]?.toLowerCase().replace(/\s+/g, "") || "unknown";
      statusBreakdown[key] = parseInt(item.Count) || 0;
    });

    // Get device inventory - CRITICAL SECTION
    const deviceData = sheets.Device_Inventory || [];
    const recentDevices = deviceData
      .filter((device) => device && device.Status === "Installed")
      .slice(0, 10)
      .map((device) => ({
        assetTag: device["Asset Tag"] || "N/A",
        userName: device["User Name"] || "Unknown",
        os: device["Operating System"] || "Unknown",
        installDate: device["Install Date"]
          ? device["Install Date"].split("T")[0]
          : "N/A",
        status: device.Status || "Unknown",
      }));

    // Get insights from Risk Assessment
    const riskData = sheets.Risk_Assessment || [];
    const insights = riskData.map((item) => ({
      type: item["Risk Level"]?.includes("High")
        ? "critical"
        : item["Risk Level"]?.includes("Medium")
        ? "warning"
        : "info",
      title: item["Risk Level"] || "Assessment",
      description: item.Description || "No description available",
      action: item.Status || "Review required",
    }));

    const totalAssets =
      parseInt(metrics.totalassets) || deviceData.length || 220;
    const installed = parseInt(metrics.installed) || 146;

    return {
      lastUpdated:
        this.excelData.metadata?.generated_at || new Date().toISOString(),
      systemStatus: "live",
      kpiMetrics: {
        totalAssets: totalAssets,
        installed: installed,
        completionRate: parseFloat(
          ((installed / totalAssets) * 100).toFixed(1)
        ),
        remaining: totalAssets - installed,
        avgPerDay: parseInt(metrics.avgperday) || 27,
        successRate: parseInt(metrics.successrate) || 133,
        responseRate: parseInt(metrics.responserate) || 72,
        issueRate: parseInt(metrics.issuerate) || 2,
        daysActive: parseInt(metrics.daysactive) || 11,
      },
      statusBreakdown: {
        installed: statusBreakdown.installed || installed,
        noResponse: statusBreakdown.noresponse || 61,
        issues: statusBreakdown.issues || 4,
        remaining: statusBreakdown.remaining || totalAssets - installed,
      },
      osDistribution: osDistribution,
      recentDevices: recentDevices,
      insights: insights,
      rawWorksheets: {
        Device_Inventory: sheets.Device_Inventory || { data: [] },
        Dashboard_Metrics: sheets.Dashboard_Metrics || { data: [] },
        OS_Distribution: sheets.OS_Distribution || { data: [] },
        Status_Breakdown: sheets.Status_Breakdown || { data: [] },
        Risk_Assessment: sheets.Risk_Assessment || { data: [] },
      },
    };
  }

  // Get empty dashboard format for fallback
  getEmptyDashboardFormat() {
    return {
      lastUpdated: new Date().toISOString(),
      systemStatus: "offline",
      kpiMetrics: {
        totalAssets: 0,
        installed: 0,
        completionRate: 0,
        remaining: 0,
        avgPerDay: 0,
        successRate: 0,
        responseRate: 0,
        issueRate: 0,
        daysActive: 0,
      },
      statusBreakdown: {
        installed: 0,
        noResponse: 0,
        issues: 0,
        remaining: 0,
      },
      osDistribution: [],
      recentDevices: [],
      insights: [],
      rawWorksheets: {
        Device_Inventory: { data: [] },
        Dashboard_Metrics: { data: [] },
        OS_Distribution: { data: [] },
        Status_Breakdown: { data: [] },
        Risk_Assessment: { data: [] },
      },
    };
  }

  // Minimal fallback data
  getFallbackData() {
    return {
      metadata: {
        source_file: "fallback",
        generated_at: new Date().toISOString(),
        total_sheets: 0,
        sheet_names: [],
      },
      sheets: {},
    };
  }

  // Populate devices table with Excel data
  populateDevicesTable() {
    const tableBody = document.getElementById("devicesTableBody");
    const deviceCountElement = document.getElementById("deviceCount");

    if (!tableBody) return;

    // Get device inventory from Excel data --> Pagination Change
    // const deviceDataMain = this.dashboardData?.rawWorksheets?.Device_Inventory?.data || [];
    const deviceDataMain =
      this.dashboardData?.rawWorksheets?.Device_Inventory || [];
    // console.log("Merry Christmas2", this.dashboardData);

    // Update device count
    if (deviceCountElement) {
      deviceCountElement.textContent = deviceDataMain.length;
    }

    this.totalDevicesCount = deviceDataMain.length;

    // Clear existing content
    tableBody.innerHTML = "";

    // Check if we have data
    if (deviceDataMain.length === 0) {
      tableBody.innerHTML = `
        <tr>
          <td colspan="6" class="empty-state">
            <div class="empty-icon">
              <i class="fas fa-database"></i>
            </div>
            <div class="empty-text">No device data available</div>
            <div class="empty-subtext">Load device inventory to view details</div>
          </td>
        </tr>`;
      return;
    }

    const start = (this.currentPage - 1) * this.recordsPerPage;
    const end = start + this.recordsPerPage;
    const deviceData = deviceDataMain.slice(start, end);

    // console.log("Debug 2", start, end);
    this.pageInfo.textContent = `Page ${this.currentPage} of ${Math.ceil(
      this.totalDevicesCount / this.recordsPerPage
    )}`;

    // Populate table with device data
    deviceData.forEach((device, index) => {
      // Skip empty rows
      if (!device || !device["Asset Tag"]) return;

      const row = document.createElement("tr");
      row.className = "device-row";

      // Determine status class and icon
      let statusClass = "unknown";
      let statusIcon = "fas fa-question-circle";
      let statusText = device.Status || "Unknown";

      if (device.Status === "Installed") {
        statusClass = "installed";
        statusIcon = "fas fa-check-circle";
      } else if (device.Notes === "No Response") {
        statusClass = "no-response";
        statusIcon = "fas fa-exclamation-triangle";
        statusText = "No Response";
      } else if (device.Notes && device.Notes.includes("Issue")) {
        statusClass = "issue";
        statusIcon = "fas fa-times-circle";
        statusText = "Issue";
      } else if (device.Notes === "Pending") {
        statusClass = "pending";
        statusIcon = "fas fa-hourglass-half";
        statusText = "Pending";
      }

      // Format installation date
      let installDate = "N/A";
      if (device["Install Date"]) {
        try {
          const date = new Date(device["Install Date"]);
          installDate = date.toLocaleDateString();
        } catch (e) {
          installDate = "N/A";
        }
      }

      // Format installation time
      let installTime = device["Installation Time"] || "N/A";
      // if (installTime && installTime !== "N/A") {
      //   // Clean up time format
      //   installTime = installTime.replace(/[-\s]/g, "").substring(0, 8);
      //   if (installTime.length >= 6) {
      //     installTime = `${installTime.substring(0, 2)}:${installTime.substring(
      //       2,
      //       4
      //     )}:${installTime.substring(4, 6)}`;
      //   }
      // }

      row.innerHTML = `
        <td >
          <div class="asset-info">
            <div class="asset-tag">${device["Asset Tag"] || "N/A"}</div>
          </div>
        </td>
            
          <td >
            <span class="user-inline">
            <i class="fas fa-user"></i>
            <span class="user-name">${
              device["User Name"] || "Unknown User"
            }</span>
          </span>
          </td>
        <td >
         
            <span class="user-inline">
              <i class="fab ${this.getOSIcon(device["Operating System"])}"></i>
            
              <span class="os-name">${
                device["Operating System"] || "Unknown"
              }</span>
            </span>
          
        </td>
        <td >
          <div class="status-badge ${statusClass}">
            <i class="${statusIcon}"></i>
            <span>${statusText}</span>
          </div>
        </td>
        <td >
          <div class="install-info">
            <!-- <div class="install-date">${installDate}</div> -->
            <div class="install-time">${installTime}</div>
          </div>
        </td>
      `;

      tableBody.appendChild(row);
    });

    // Add click handlers for device actions
    this.addDeviceActionHandlers();
  }

  // Get appropriate OS icon class
  getOSIcon(os) {
    if (!os) return "fa-question-circle";

    const osLower = os.toLowerCase();
    if (osLower.includes("windows")) return "fa-windows";
    if (osLower.includes("mac") || osLower.includes("darwin"))
      return "fa-apple";
    if (osLower.includes("linux")) return "fa-linux";
    return "fa-desktop";
  }

  // Add device action handlers
  addDeviceActionHandlers() {
    const actionButtons = document.querySelectorAll(
      ".device-actions .action-btn"
    );

    actionButtons.forEach((button) => {
      button.addEventListener("click", (e) => {
        e.stopPropagation();
        const deviceIndex = parseInt(button.getAttribute("data-device-index"));
        const action = button.classList.contains("view-details")
          ? "view"
          : "refresh";

        this.handleDeviceAction(action, deviceIndex);
      });
    });
  }

  // Handle device actions
  handleDeviceAction(action, deviceIndex) {
    const deviceData =
      this.dashboardData?.rawWorksheets?.Device_Inventory?.data || [];
    const device = deviceData[deviceIndex];

    if (!device) return;

    if (action === "view") {
      this.showDeviceDetails(device);
    } else if (action === "refresh") {
      this.refreshDeviceStatus(device, deviceIndex);
    }
  }

  // Show device details modal
  showDeviceDetails(device) {
    const modal = document.createElement("div");
    modal.className = "device-modal-overlay";

    modal.innerHTML = `
      <div class="device-modal">
        <div class="modal-header">
          <h3><i class="fas fa-laptop-code"></i> Device Details</h3>
          <button class="modal-close">&times;</button>
        </div>
        <div class="modal-content">
          <div class="device-detail-grid">
            <div class="detail-item">
              <label>Asset Tag</label>
              <span>${device["Asset Tag"] || "N/A"}</span>
            </div>
            <div class="detail-item">
              <label>User Name</label>
              <span>${device["User Name"] || "Unknown"}</span>
            </div>
            <div class="detail-item">
              <label>Operating System</label>
              <span>${device["Operating System"] || "Unknown"}</span>
            </div>
            <div class="detail-item">
              <label>Status</label>
              <span class="status-value">${device.Status || "Unknown"}</span>
            </div>
            <div class="detail-item">
              <label>Install Date</label>
              <span>${
                device["Install Date"]
                  ? new Date(device["Install Date"]).toLocaleString()
                  : "N/A"
              }</span>
            </div>
            <div class="detail-item">
              <label>Installation Time</label>
              <span>${device["Installation Time"] || "N/A"}</span>
            </div>
            <div class="detail-item full-width">
              <label>Notes</label>
              <span>${device.Notes || "No notes available"}</span>
            </div>
          </div>
        </div>
      </div>
    `;

    document.body.appendChild(modal);

    // Add close handlers
    modal.querySelector(".modal-close").addEventListener("click", () => {
      modal.remove();
    });

    modal.addEventListener("click", (e) => {
      if (e.target === modal) {
        modal.remove();
      }
    });
  }

  // Refresh device status
  refreshDeviceStatus(device, deviceIndex) {
    this.showNotification(
      `Refreshing status for ${device["Asset Tag"]}...`,
      "info",
      2000
    );

    // Simulate refresh (in real implementation, this would make an API call)
    setTimeout(() => {
      this.showNotification(
        `Status updated for ${device["Asset Tag"]}`,
        "success",
        3000
      );
    }, 1500);
  }

  // Setup event listeners
  setupEventListeners() {
    // Tab navigation
    const tabButtons = document.querySelectorAll(".live-tab-btn");
    const tabPanels = document.querySelectorAll(".tab-panel");

    tabButtons.forEach((button, index) => {
      button.addEventListener("click", () => {
        // Remove active class from all tabs and panels
        tabButtons.forEach((btn) => btn.classList.remove("active"));
        tabPanels.forEach((panel) => panel.classList.remove("active"));

        // Add active class to clicked tab and corresponding panel
        button.classList.add("active");
        const tabId = button.getAttribute("data-tab");
        const panel = document.getElementById(tabId);
        if (panel) {
          panel.classList.add("active");
        }

        // Update tab slider position
        this.updateTabSlider(index);
      });
    });

    // Refresh and export buttons
    const refreshBtn = document.querySelector(".control-btn.refresh");
    const exportBtn = document.querySelector(".control-btn.export");

    if (refreshBtn) {
      refreshBtn.addEventListener("click", () => {
        this.loadDashboardData();
      });
    }

    if (exportBtn) {
      exportBtn.addEventListener("click", () => {
        this.exportDeviceReport();
      });
    }

    // Device search
    const searchInput = document.getElementById("deviceSearch");
    if (searchInput) {
      searchInput.addEventListener("input", (e) => {
        this.filterDevices(e.target.value);
      });
    }
  }

  // Update tab slider position
  updateTabSlider(activeIndex) {
    const slider = document.querySelector(".tab-slider");
    const buttons = document.querySelectorAll(".live-tab-btn");

    if (slider && buttons[activeIndex]) {
      const activeButton = buttons[activeIndex];
      const sliderWidth = activeButton.offsetWidth;
      const sliderLeft = activeButton.offsetLeft;

      slider.style.width = sliderWidth + "px";
      slider.style.left = sliderLeft + "px";
    }
  }

  // Initialize tooltips
  initializeTooltips() {
    const tooltipElements = document.querySelectorAll("[data-tooltip]");

    tooltipElements.forEach((element) => {
      element.addEventListener("mouseenter", (e) => {
        this.showTooltip(e.target, e.target.getAttribute("data-tooltip"));
      });

      element.addEventListener("mouseleave", () => {
        this.hideTooltip();
      });
    });
  }

  // Show tooltip
  showTooltip(element, text) {
    let tooltip = document.getElementById("tooltip");
    if (!tooltip) {
      tooltip = document.createElement("div");
      tooltip.id = "tooltip";
      tooltip.className = "tooltip";
      document.body.appendChild(tooltip);
    }

    tooltip.innerHTML = `
      <div class="tooltip-content">${text}</div>
      <div class="tooltip-arrow"></div>
    `;

    const rect = element.getBoundingClientRect();
    tooltip.style.left =
      rect.left + rect.width / 2 - tooltip.offsetWidth / 2 + "px";
    tooltip.style.top = rect.top - tooltip.offsetHeight - 10 + "px";
    tooltip.style.opacity = "1";
    tooltip.style.visibility = "visible";
  }

  // Hide tooltip
  hideTooltip() {
    const tooltip = document.getElementById("tooltip");
    if (tooltip) {
      tooltip.style.opacity = "0";
      tooltip.style.visibility = "hidden";
    }
  }

  // Initialize live clock
  initializeLiveClock() {
    const updateClock = () => {
      const now = new Date();
      const timeString = now.toLocaleTimeString("en-US", {
        hour12: false,
        hour: "2-digit",
        minute: "2-digit",
        second: "2-digit",
      });

      const clockElement = document.getElementById("liveClock");
      if (clockElement) {
        clockElement.textContent = timeString;
      }

      // Update last sync time
      const lastSyncElement = document.getElementById("lastSyncTime");
      if (lastSyncElement && this.dashboardData) {
        const lastUpdated = new Date(this.dashboardData.lastUpdated);
        const timeDiff = Math.floor((now - lastUpdated) / 1000);

        let syncText = "Just now";
        if (timeDiff > 60) {
          syncText = `${Math.floor(timeDiff / 60)}m ago`;
        } else if (timeDiff > 0) {
          syncText = `${timeDiff}s ago`;
        }

        lastSyncElement.textContent = syncText;
      }
    };

    updateClock();
    setInterval(updateClock, 1000);
  }

  // Update dashboard with Excel data
  async updateDashboard() {
    if (!this.dashboardData) return;

    try {
      await Promise.all([
        this.updateKPICards(),
        this.updateProgressCards(),
        this.updateStatsCards(),
        this.populateDevicesTable(),
        this.updateAnalytics(),
        this.renderOSChart(),
        this.updateInsights(),
      ]);
    } catch (error) {
      console.error("Error updating dashboard:", error);
    }
  }

  updateAnalytics() {
    document.getElementById("os_distribution_analytics").innerHTML = `
    <div class="os-stat-item windows11">
                <div class="os-icon-wrapper">
                  <i class="fab fa-windows"></i>
                  <div class="version-badge">10</div>
                </div>
                <div class="os-info">
                  <!-- <div class="os-count" data-target="59">0</div> -->
                  <div class="os-name">${this.dashboardData.osDistribution[0].name}</div>
                  <div class="os-details">
                    <span class="percentage">${this.dashboardData.osDistribution[0].percentage}%</span>
                    <span class="version">${this.dashboardData.osDistribution[0].version}</span>
                  </div>
                </div>
                <div class="os-progress-ring">
                  <svg width="60" height="60">
                    <circle
                      cx="30"
                      cy="30"
                      r="25"
                      stroke="rgba(59, 130, 246, 0.2)"
                      stroke-width="4"
                      fill="none"
                    />
                    <circle
                      cx="30"
                      cy="30"
                      r="25"
                      stroke="#3b82f6"
                      stroke-width="4"
                      fill="none"
                      stroke-dasharray="157.08"
                      stroke-dashoffset="12.25"
                      stroke-linecap="round"
                    />
                  </svg>
                  <div class="ring-percentage">${this.dashboardData.osDistribution[0].percentage}%</div>
                </div>
              </div>

              <div class="os-stat-item windows10">
                <div class="os-icon-wrapper">
                  <i class="fab fa-windows"></i>
                  <div class="version-badge">11</div>
                </div>
                <div class="os-info">
                  <!-- <div class="os-count" data-target="4">0</div> -->
                  <div class="os-name">${this.dashboardData.osDistribution[1].name}</div>
                  <div class="os-details">
                    <span class="percentage">${this.dashboardData.osDistribution[1].percentage}%</span>
                    <span class="version">${this.dashboardData.osDistribution[1].version}</span>
                  </div>
                </div>
                <div class="os-progress-ring">
                  <svg width="60" height="60">
                    <circle
                      cx="30"
                      cy="30"
                      r="25"
                      stroke="rgba(245, 158, 11, 0.2)"
                      stroke-width="4"
                      fill="none"
                    />
                    <circle
                      cx="30"
                      cy="30"
                      r="25"
                      stroke="#f59e0b"
                      stroke-width="4"
                      fill="none"
                      stroke-dasharray="157.08"
                      stroke-dashoffset="147.34"
                      stroke-linecap="round"
                    />
                  </svg>
                  <div class="ring-percentage">${this.dashboardData.osDistribution[1].percentage}%</div>
                </div>
              </div>

              <div class="os-stat-item macos">
                <div class="os-icon-wrapper">
                  <i class="fab fa-apple"></i>
                  <div class="version-badge">14</div>
                </div>
                <div class="os-info">
                  <!-- <div class="os-count" data-target="1">0</div> -->
                  <div class="os-name">${this.dashboardData.osDistribution[2].name}</div>
                  <div class="os-details">
                    <span class="percentage">${this.dashboardData.osDistribution[2].percentage}%</span>
                    <span class="version">${this.dashboardData.osDistribution[2].version}</span>
                  </div>
                </div>
                <div class="os-progress-ring">
                  <svg width="60" height="60">
                    <circle
                      cx="30"
                      cy="30"
                      r="25"
                      stroke="rgba(148, 163, 184, 0.2)"
                      stroke-width="4"
                      fill="none"
                    />
                    <circle
                      cx="30"
                      cy="30"
                      r="25"
                      stroke="#94a3b8"
                      stroke-width="4"
                      fill="none"
                      stroke-dasharray="157.08"
                      stroke-dashoffset="154.57"
                      stroke-linecap="round"
                    />
                  </svg>
                  <div class="ring-percentage">${this.dashboardData.osDistribution[2].percentage}%</div>
                </div>
              </div>
    `;

    document.getElementById("os-distribution-bar-chart").innerHTML = `
    <div class="os-chart-item" data-os="windows11">
                <div class="os-label">
                  <i class="fab fa-windows"></i>
                  <span class="os-name">${this.dashboardData.osDistribution[0].name}</span>
                  <span class="os-version">${this.dashboardData.osDistribution[0].version}</span>
                </div>
                <div class="os-bar-container">
                  <div class="os-bar" data-percentage="84.6">
                    <div class="os-bar-fill" style="width: 84.6%"></div>
                    <div class="os-bar-shimmer"></div>
                  </div>
                  <div class="os-count">${this.dashboardData.osDistribution[0].value} devices</div>
                </div>
                <div class="os-percentage">${this.dashboardData.osDistribution[0].percentage}%</div>
              </div>
          
              <div class="os-chart-item" data-os="windows10">
                <div class="os-label">
                  <i class="fab fa-windows"></i>
                  <span class="os-name">${this.dashboardData.osDistribution[1].name}</span>
                  <span class="os-version">${this.dashboardData.osDistribution[1].version}</span>
                </div>
                <div class="os-bar-container">
                  <div class="os-bar" data-percentage="13.4">
                    <div class="os-bar-fill" style="width: 13.4%"></div>
                    <div class="os-bar-shimmer"></div>
                  </div>
                  <div class="os-count">${this.dashboardData.osDistribution[1].value} devices</div>
                </div>
                <div class="os-percentage">${this.dashboardData.osDistribution[1].percentage}%</div>
              </div>
 
              <div class="os-chart-item" data-os="macos">
                <div class="os-label">
                  <i class="fab fa-apple"></i>
                  <span class="os-name">${this.dashboardData.osDistribution[2].name}</span>
                  <span class="os-version">${this.dashboardData.osDistribution[2].version}</span>
                </div>
                <div class="os-bar-container">
                  <div class="os-bar" data-percentage="2.0">
                    <div class="os-bar-fill" style="width: 2.0%"></div>
                    <div class="os-bar-shimmer"></div>
                  </div>
                  <div class="os-count">${this.dashboardData.osDistribution[2].value} devices</div>
                </div>
                <div class="os-percentage">${this.dashboardData.osDistribution[2].percentage}%</div>
              </div>
    `;
  }

  // Update KPI cards with Excel metrics
  updateKPICards() {
    const kpiData = this.dashboardData.kpiMetrics;

    // Update completion rate
    const completionElement = document.querySelector('[data-target="28.8"]');
    if (completionElement) {
      completionElement.setAttribute("data-target", kpiData.completionRate);
      completionElement.textContent = kpiData.completionRate + "%";
    }

    // Update installed count
    const installedElement = document.querySelector('[data-target="64"]');
    if (installedElement) {
      installedElement.setAttribute("data-target", kpiData.installed);
      this.animateCounter(installedElement, 0, kpiData.installed, 2000);
    }

    // Update no response count
    const noResponseElement = document.querySelector('[data-target="5"]');
    if (noResponseElement) {
      const noResponseCount = this.dashboardData.statusBreakdown.noResponse;
      noResponseElement.setAttribute("data-target", noResponseCount);
      this.animateCounter(noResponseElement, 0, noResponseCount, 2000);
    }

    // Update issues count
    const issuesElement = document.querySelector('[data-target="1"]');
    if (issuesElement) {
      const issuesCount = this.dashboardData.statusBreakdown.issues;
      issuesElement.setAttribute("data-target", issuesCount);
      this.animateCounter(issuesElement, 0, issuesCount, 2000);
    }
  }

  // Update progress cards
  updateProgressCards() {
    const kpiData = this.dashboardData.kpiMetrics;

    // Update completed progress
    const completedBar = document.querySelector(".completed-fill");
    if (completedBar) {
      completedBar.style.width = kpiData.completionRate + "%";
    }

    const completedPercentage = document.querySelector(
      ".progress-card.completed .progress-percentage"
    );
    if (completedPercentage) {
      completedPercentage.textContent = kpiData.completionRate + "%";
    }

    const completedValue = document.querySelector(
      ".progress-card.completed .progress-value"
    );
    if (completedValue) {
      completedValue.textContent = `${kpiData.installed} / ${kpiData.totalAssets}`;
    }

    // Update remaining progress
    const remainingRate = 100 - kpiData.completionRate;
    const remainingBar = document.querySelector(".remaining-fill");
    if (remainingBar) {
      remainingBar.style.width = remainingRate + "%";
    }

    const remainingPercentage = document.querySelector(
      ".progress-card.remaining .progress-percentage"
    );
    if (remainingPercentage) {
      remainingPercentage.textContent = remainingRate.toFixed(1) + "%";
    }

    const remainingValue = document.querySelector(
      ".progress-card.remaining .progress-value"
    );
    if (remainingValue) {
      remainingValue.textContent = `${kpiData.remaining} / ${kpiData.totalAssets}`;
    }
  }

  // Update stats cards
  updateStatsCards() {
    const kpiData = this.dashboardData.kpiMetrics;

    // Total assets
    const totalAssetsElement = document.querySelector('[data-target="222"]');
    if (totalAssetsElement) {
      this.animateCounter(totalAssetsElement, 0, kpiData.totalAssets, 2000);
    }

    // Protected (installed)
    const protectedElement = document.querySelector(
      '[data-target="64"]:not(.live-kpi-card *)'
    );
    if (protectedElement) {
      this.animateCounter(protectedElement, 0, kpiData.installed, 2000);
    }

    // Remaining
    const remainingElement = document.querySelector('[data-target="158"]');
    if (remainingElement) {
      this.animateCounter(remainingElement, 0, kpiData.remaining, 2000);
    }

    // Days active
    const daysActiveElement = document.querySelector('[data-target="9"]');
    if (daysActiveElement) {
      this.animateCounter(daysActiveElement, 0, kpiData.daysActive, 2000);
    }

    // Success rate
    const successRateElement = document.querySelector('[data-target="63.4"]');
    if (successRateElement) {
      successRateElement.textContent = kpiData.successRate.toFixed(1);
    }
  }

  // Animate counter
  animateCounter(element, start, end, duration) {
    const startTime = performance.now();

    const animate = (currentTime) => {
      const elapsed = currentTime - startTime;
      const progress = Math.min(elapsed / duration, 1);

      const current = Math.floor(start + (end - start) * progress);
      element.textContent = current;

      if (progress < 1) {
        requestAnimationFrame(animate);
      }
    };

    requestAnimationFrame(animate);
  }

  // Render OS Chart with Excel data
  renderOSChart() {
    const canvas = document.getElementById("osChart");
    if (!canvas) return;

    const ctx = canvas.getContext("2d");

    if (this.osChart) {
      this.osChart.destroy();
    }

    const osData = this.dashboardData.osDistribution;

    this.osChart = new Chart(ctx, {
      type: "doughnut",
      data: {
        labels: osData.map((item) => item.name),
        datasets: [
          {
            data: osData.map((item) => item.value),
            backgroundColor: [
              this.chartColors.primary,
              this.chartColors.warning,
              this.chartColors.info,
            ],
            borderColor: [
              this.chartBorderColors.primary,
              this.chartBorderColors.warning,
              this.chartBorderColors.info,
            ],
            borderWidth: 2,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            display: false,
          },
        },
      },
    });
  }

  // Update insights with Risk Assessment data
  updateInsights() {
    // This would update the action items based on Risk_Assessment sheet
    // Implementation depends on specific UI elements in the actions tab
  }

  // Show loading overlay
  showLoadingOverlay() {
    const overlay = document.querySelector(".loading-overlay");
    if (overlay) {
      overlay.style.display = "flex";
    }
  }

  // Hide loading overlay
  hideLoadingOverlay() {
    const overlay = document.querySelector(".loading-overlay");
    if (overlay) {
      overlay.style.display = "none";
    }
  }

  // Show notification
  showNotification(message, type = "info", duration = 3000) {
    const notification = document.createElement("div");
    notification.className = `notification ${type}`;
    notification.innerHTML = `
      <i class="fas ${
        type === "success"
          ? "fa-check-circle"
          : type === "warning"
          ? "fa-exclamation-triangle"
          : "fa-info-circle"
      }"></i>
      <span>${message}</span>
    `;

    document.body.appendChild(notification);

    // Show notification
    setTimeout(() => {
      notification.classList.add("show");
    }, 100);

    // Hide and remove notification
    setTimeout(() => {
      notification.classList.remove("show");
      setTimeout(() => {
        document.body.removeChild(notification);
      }, 300);
    }, duration);
  }

  // Filter devices based on search
  filterDevices(searchTerm) {
    const rows = document.querySelectorAll("#devicesTableBody .device-row");

    rows.forEach((row) => {
      const text = row.textContent.toLowerCase();
      const matches = text.includes(searchTerm.toLowerCase());
      row.style.display = matches ? "" : "none";
    });
  }

  // Export device report
  exportDeviceReport() {
    const deviceData =
      this.dashboardData?.rawWorksheets?.Device_Inventory?.data || [];

    if (deviceData.length === 0) {
      this.showNotification("No device data to export", "warning", 3000);
      return;
    }

    // Create CSV content
    const headers = [
      "Asset Tag",
      "User Name",
      "Operating System",
      "Status",
      "Install Date",
      "Notes",
    ];
    const csvContent = [
      headers.join(","),
      ...deviceData.map((device) =>
        [
          `"${device["Asset Tag"] || ""}"`,
          `"${device["User Name"] || ""}"`,
          `"${device["Operating System"] || ""}"`,
          `"${device.Status || ""}"`,
          `"${device["Install Date"] || ""}"`,
          `"${device.Notes || ""}"`,
        ].join(",")
      ),
    ].join("\n");

    // Download CSV
    const blob = new Blob([csvContent], { type: "text/csv" });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "crowdstrike-device-report.csv";
    a.click();
    window.URL.revokeObjectURL(url);

    this.showNotification(
      "Device report exported successfully",
      "success",
      3000
    );
  }
} //end of the class

// Initialize dashboard when DOM is loaded
document.addEventListener("DOMContentLoaded", () => {
  window.dashboard = new SuperLiveDashboard();
});

// Refresh devices function (global)
function refreshDevices() {
  if (window.dashboard) {
    window.dashboard.loadDashboardData();
  }
}

// Print devices function (global)
function printDevices() {
  if (window.dashboard) {
    window.dashboard.exportDeviceReport();
  }
}

function paginatorForward() {
  window.dashboard.paginatorNxtBtn();
}

function paginatorBackward() {
  window.dashboard.paginatorPreBtn();
}

const paginatorBtn = document.getElementById("recordsPerPage");
paginatorBtn.addEventListener("change", (e) => {
  window.dashboard.setPaginatorElements(1, e.target.value);
  window.dashboard.populateDevicesTable();
});
