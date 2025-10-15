

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

    console.log("Happy Christmas");


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
      const response = await fetch("./dashboard_data.json");
      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      this.excelData = await response.json();

      // this.excelData = {
      //   metadata: {
      //     source_file:
      //       "CrowdStrike-EDR-Installation-Status-v1.0-20251007-1.xlsx",
      //     generated_at: "2025-10-13T05:56:57.661357",
      //     total_sheets: 5,
      //     sheet_names: [
      //       "Device_Inventory",
      //       "Dashboard_Metrics",
      //       "OS_Distribution",
      //       "Status_Breakdown",
      //       "Risk_Assessment",
      //     ],
      //   },
      //   sheets: {
      //     Device_Inventory: {
      //       row_count: 221,
      //       column_count: 7,
      //       columns: [
      //         "Asset Tag",
      //         "User Name",
      //         "Operating System",
      //         "Status",
      //         "Installation Time",
      //         "Install Date",
      //         "Notes",
      //       ],
      //       data: [
      //         {
      //           "Asset Tag": "LAP 898",
      //           "User Name": "Panduka Pathirathna",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-04:31:00",
      //         },
      //         {
      //           "Asset Tag": "LAP 895",
      //           "User Name": "Navindu Vidanagama",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-07:10:00",
      //           "Install Date": "2025-10-07T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 829",
      //           "User Name": "Nipuni Kodithuwakku",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-04:28:97",
      //           "Install Date": "2025-10-03T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 856",
      //           "User Name": "Ravindu Dilshan",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-04:32:00 ",
      //           "Install Date": "2025-10-03T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 846",
      //           "User Name": "Mohamed Minsaf",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-04:28:13",
      //           "Install Date": "2025-10-08T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 710 ",
      //           "User Name": "Senerath Alahakon",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-03:47:23",
      //           "Install Date": "2025-10-07T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 848",
      //           "User Name": "Kasun Wanniarachchi",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-08:18:44",
      //           "Install Date": "2025-10-10T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 590",
      //           "User Name": "Indrakumara Sirisena",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-09:38:20",
      //           "Install Date": "2025-10-07T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 876",
      //           "User Name": "Madusha Weerasinghe",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-10:32:00",
      //           "Install Date": "2025-10-07T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 781",
      //           "User Name": "Lahiru Bandara",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 760",
      //           "User Name": "Hiruni Jayasinghe",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-08:18:44",
      //           "Install Date": "2025-10-03T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 839",
      //           "User Name": "Suneth Ekanayaka",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-05:20:00",
      //           "Install Date": "2025-10-03T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 786",
      //           "User Name": "Danushka Bandara",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-03:49:06",
      //           "Install Date": "2025-10-03T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 767",
      //           "User Name": "Jeema Riyana",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-04:51:64",
      //           "Install Date": "2025-10-03T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 858",
      //           "User Name": "Asanka Kumara",
      //           "Operating System": "Windows 10",
      //           Status: "Installed",
      //           "Installation Time": "-04:31:00",
      //           "Install Date": "2025-10-03T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 778",
      //           "User Name": "Ishara Jayasinghe",
      //           "Operating System": "Windows 10",
      //           Status: "Installed",
      //           "Installation Time": "-04:02:74",
      //           "Install Date": "2025-10-03T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 889",
      //           "User Name": "Tharanga Liyanage",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-03:14:00",
      //           "Install Date": "2025-10-03T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 901",
      //           "User Name": "Ayesh Anushanga",
      //           "Operating System": "macOS",
      //           Status: "Installed",
      //           "Installation Time": "-06:32:40",
      //           "Install Date": "2025-10-10T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 795",
      //           "User Name": "Inusha Udakanjalee",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-13:28:19",
      //           "Install Date": "2025-10-03T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 873",
      //           "User Name": "Piumi Dinuka",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-04:31:97\t",
      //           "Install Date": "2025-10-03T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 897",
      //           "User Name": "Vishmantha Poramba Badalge",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-04:05:00",
      //           "Install Date": "2025-10-03T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 680",
      //           "User Name": "Sachith Kaushalya",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-03:41:87",
      //           "Install Date": "2025-10-03T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 675",
      //           "User Name": "Ravindu Chinthaka",
      //           "Operating System": "Windows 10",
      //           Status: "Installed",
      //           "Installation Time": "-05:09:00",
      //           "Install Date": "2025-10-03T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 859",
      //           "User Name": "Muhannadh Razick",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-03:27:00",
      //           "Install Date": "2025-10-03T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 793",
      //           "User Name": "Tharindu Kumara",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-06:20:00",
      //           "Install Date": "2025-10-03T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 729",
      //           "User Name": "Gayalan Kishor",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-06:08:00",
      //           "Install Date": "2025-10-03T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 866",
      //           "User Name": "Jathishwarya Venugopalavanit",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-04:25:00",
      //           "Install Date": "2025-10-03T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 661",
      //           "User Name": "Piyumitha Nirman",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-03:41:79",
      //           "Install Date": "2025-10-03T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 893",
      //           "User Name": "Isara Tillekeratne",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-03:39:00",
      //           "Install Date": "2025-10-03T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "TAPRO LAP 029",
      //           "User Name": "Kanishka Perera",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-05:07:00",
      //           "Install Date": "2025-10-07T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 863",
      //           "User Name": "Supun Adikari",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": " 03:13:00",
      //           "Install Date": "2025-10-07T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 833",
      //           "User Name": "Tharindu disanayaka",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": " 04:35:00",
      //           "Install Date": "2025-10-07T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 640",
      //           "User Name": "Imesh Fernando",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-03:41:87",
      //           "Install Date": "2025-10-08T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 685",
      //           "User Name": "Varuni Punchihewa",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-08:27:03",
      //           "Install Date": "2025-10-07T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 782",
      //           "User Name": "Rashan Udayanka",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-04:48:09",
      //           "Install Date": "2025-10-07T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 882",
      //           "User Name": "Maheshika Wickramasinghe",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-10:45:33",
      //           "Install Date": "2025-10-07T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 687",
      //           "User Name": "Induwara Rupasinghe",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-05:43:50",
      //           "Install Date": "2025-10-07T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 738",
      //           "User Name": "Supuni Wijerathne",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-07:56:18",
      //           "Install Date": "2025-10-07T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 699",
      //           "User Name": "Emindu Perera",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-05:18:44",
      //           "Install Date": "2025-10-07T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 892",
      //           "User Name": "Anuradhal",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-04:07:64",
      //           "Install Date": "2025-10-07T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 894",
      //           "User Name": "Wathmi Thennakoon",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-05:46:86",
      //           "Install Date": "2025-10-07T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 853",
      //           "User Name": "Masha Pupulewatte",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-04:28:00",
      //           "Install Date": "2025-10-08T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 742",
      //           "User Name": "Janudha Gunawardena",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-04:56:52",
      //           "Install Date": "2025-10-07T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 908",
      //           "User Name": "Jayalal Kahandawa",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-06:59:31",
      //           "Install Date": "2025-10-07T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 716",
      //           "User Name": "Malidu Jayasundara",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-05:50:74",
      //           "Install Date": "2025-10-08T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 593",
      //           "User Name": "Jayamal Gunawardhana",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-03:51:63",
      //           "Install Date": "2025-10-08T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 635",
      //           "User Name": "Dushantha Mathes",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-04:51:89",
      //           "Install Date": "2025-10-08T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 727",
      //           "User Name": "Priyanke Wijesekara",
      //           "Operating System": "Windows 10",
      //           Status: "Installed",
      //           "Installation Time": "-07:08:19",
      //           "Install Date": "2025-10-10T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 747",
      //           "User Name": "Ahmed Shafraz",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-05:53:30",
      //           "Install Date": "2025-10-08T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 682",
      //           "User Name": "Shakya Wanigarathna",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-05:40:53",
      //           "Install Date": "2025-10-08T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 641",
      //           "User Name": "Dulanka Karunasena",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-04:26:51",
      //           "Install Date": "2025-10-08T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 840",
      //           "User Name": "Dasun Manathunga",
      //           "Operating System": "Windows 10",
      //           Status: "Installed",
      //           "Installation Time": "-05:23:54",
      //           "Install Date": "2025-10-08T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 880",
      //           "User Name": "Hirumi Kodithuwakku",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-04:09:44",
      //           "Install Date": "2025-10-08T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 877",
      //           "User Name": "Chathushag",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-06:27:55",
      //           "Install Date": "2025-10-08T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 579",
      //           "User Name": "Kusal Kahaduwa",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-05:28:59",
      //           "Install Date": "2025-10-08T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 806",
      //           "User Name": "Geeth Tharanga",
      //           "Operating System": "Windows 10",
      //           Status: "Installed",
      //           "Installation Time": "-05:23:54",
      //           "Install Date": "2025-10-08T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 548",
      //           "User Name": "Yasiru Pathirana",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-11:08:17",
      //           "Install Date": "2025-10-08T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 705",
      //           "User Name": "Vimukthi Mayadunne",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-06:25:44",
      //           "Install Date": "2025-10-08T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 684",
      //           "User Name": "Thilina Jayamini",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-05:59:12",
      //           "Install Date": "2025-10-08T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 841",
      //           "User Name": "Yasas Sandeepa",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-05:49:02",
      //           "Install Date": "2025-10-08T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 888",
      //           "User Name": "Mohommad Shamil",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-05:59:12",
      //           "Install Date": "2025-10-08T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 852",
      //           "User Name": "Osura Hettiarachchi",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-20:02:25",
      //           "Install Date": "2025-10-09T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 902",
      //           "User Name": "Osura Hettiarachchi",
      //           "Operating System": "macOS",
      //           Status: "Installed",
      //           "Installation Time": "-05:00:00",
      //           "Install Date": "2025-10-09T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 899 ",
      //           "User Name": "Shanupa Liyanaarachchi",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-07:27:03",
      //           "Install Date": "2025-10-09T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 664",
      //           "User Name": "Geethika Kalhari",
      //           "Operating System": "Windows 10",
      //           Status: "Installed",
      //           "Installation Time": "-05:17:26",
      //           "Install Date": "2025-10-10T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 891",
      //           "User Name": "Devindi Jayathilaka",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-06:03:42",
      //           "Install Date": "2025-10-10T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 886",
      //           "User Name": "Sumeera Liyanage",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-05:56:03",
      //           "Install Date": "2025-10-10T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 855",
      //           "User Name": "Thushan Lakshitha",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-09:24:10",
      //           "Install Date": "2025-10-09T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 804",
      //           "User Name": "Lalinda Udurawana",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-04:37:11",
      //           "Install Date": "2025-10-09T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 752",
      //           "User Name": "Raveesha Fernando",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-02:24:52",
      //           "Install Date": "2025-10-09T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 854",
      //           "User Name": "Yasara Jayadinee",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-02:45:03",
      //           "Install Date": "2025-10-09T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 810",
      //           "User Name": "Kalana Gunathilaka",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-05:49:39",
      //           "Install Date": "2025-10-09T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 865",
      //           "User Name": "Sajeepan Srithararasan",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-06:15:57",
      //           "Install Date": "2025-10-10T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 861",
      //           "User Name": "Diwanga Amasith",
      //           "Operating System": "Windows 10",
      //           Status: "Installed",
      //           "Installation Time": "-06:44:10",
      //           "Install Date": "2025-10-10T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 835",
      //           "User Name": "Janindu Ranaweera",
      //           "Operating System": "Windows 10",
      //           Status: "Installed",
      //           "Installation Time": "-06:36:52",
      //           "Install Date": "2025-10-10T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 733",
      //           "User Name": "Hashini Siriwardhane",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-07:41:33",
      //           "Install Date": "2025-10-10T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 845",
      //           "User Name": "Lakshan Chathuranga",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-05:52:11",
      //           "Install Date": "2025-10-10T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 887",
      //           "User Name": "Ashfa Assath",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-05:09:21",
      //           "Install Date": "2025-10-10T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 751",
      //           "User Name": "Hiruni Kuruppu",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 857",
      //           "User Name": "Upeksha Udayaratne",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-06:14:58",
      //           "Install Date": "2025-10-10T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 805",
      //           "User Name": "yesitha",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-05:21:13",
      //           "Install Date": "2025-10-10T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 487",
      //           "User Name": "Danusha Navod",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-06:54:10",
      //           "Install Date": "2025-10-10T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 881",
      //           "User Name": "Kavindu Rathnasekara",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-05:28:50",
      //           "Install Date": "2025-10-10T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 885",
      //           "User Name": "Achala Fernando",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-05:07:26",
      //           "Install Date": "2025-10-10T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 851",
      //           "User Name": "Sanduni Hansika",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-06:18:05",
      //           "Install Date": "2025-10-10T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 788",
      //           "User Name": "Shamila De Silva",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-08:08:17",
      //           "Install Date": "2025-10-10T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 906",
      //           "User Name": "Viduranga Randila ",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 726",
      //           "User Name": "Satheeq Hassan",
      //           "Operating System": "macOS",
      //           Status: "Installed",
      //           "Installation Time": "-05:41:20",
      //           "Install Date": "2025-10-10T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 872",
      //           "User Name": "Punsiru Alwis",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "Pending",
      //         },
      //         {
      //           "Asset Tag": "LAP 551",
      //           "User Name": "Chathuranga Mohottala",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 609 ",
      //           "User Name": "Chathuranga Mohottala",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 875",
      //           "User Name": "Suranga Caldera",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 741",
      //           "User Name": "Wathsala Wijekoon",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "Pending",
      //         },
      //         {
      //           "Asset Tag": "LAP 867",
      //           "User Name": "Sajith Perera",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 811",
      //           "User Name": "Adeesha Peiris",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 832",
      //           "User Name": "Fahim Feroz",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 746",
      //           "User Name": "Anolie Kumarasinghe",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-06:44:10",
      //           "Install Date": "2025-10-10T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 850",
      //           "User Name": "Chilanka Halpage",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 557",
      //           "User Name": "Yohan Ranasinghe",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 799",
      //           "User Name": "Pathumi Tharuka",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-06:57:29",
      //           "Install Date": "2025-10-10T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 674",
      //           "User Name": "Mohomed Rimnaz",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 779",
      //           "User Name": "Chanaka De Silva",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 731",
      //           "User Name": "Menuka Perera ",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 787",
      //           "User Name": "Isuru Sampath",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 860",
      //           "User Name": "Gayan Gunarathne",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 884",
      //           "User Name": "Prageeth Wimalarathna",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "- 09:00:00",
      //           "Install Date": "2025-10-07T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": null,
      //           "User Name": null,
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: null,
      //         },
      //         {
      //           "Asset Tag": "TAPRO LAP 016",
      //           "User Name": "Nikini Panagoda",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": null,
      //           "Install Date": "2025-10-02T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 665",
      //           "User Name": "Jeyakumar Manoj",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": null,
      //           "Install Date": "2025-10-02T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "TAPRO LAP 059",
      //           "User Name": "Hasantha Pathirana",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": null,
      //           "Install Date": "2025-10-03T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "TAPRO LAP 052",
      //           "User Name": "Kalana Abeysiriwardhana",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "- 04:18:00",
      //           "Install Date": "2025-10-03T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 871",
      //           "User Name": "Sanjeeva Samarawira",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "- 02:40:00",
      //           "Install Date": "2025-10-03T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "TAPRO LAP 050",
      //           "User Name": "Hasala Vithanage",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "- 03:15:00",
      //           "Install Date": "2025-10-03T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "TAPRO LAP 037",
      //           "User Name": "Isuru Senanayake",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "- 03:17:00",
      //           "Install Date": "2025-10-03T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "TAPRO LAP 039",
      //           "User Name": "Naflan Thowfeek",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-04:06:00",
      //           "Install Date": "2025-10-07T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "TAPRO LAP 023",
      //           "User Name": "Hansika Sadaruwani",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 757",
      //           "User Name": "Shanika Perera",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": null,
      //           "Install Date": "2025-07-01T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 808",
      //           "User Name": "Hasitha Edirisinghe",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "- 02:47:00",
      //           "Install Date": "2025-10-07T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 724",
      //           "User Name": "Prageeth Liyanage",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "- 02:55:00",
      //           "Install Date": "2025-10-07T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 803",
      //           "User Name": "Champaka Dammage",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-04:03:00",
      //           "Install Date": "2025-10-07T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 643",
      //           "User Name": "Ridmi Pinsirini",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "- 04:19:00",
      //           "Install Date": "2025-10-07T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 547",
      //           "User Name": "Gayan Ranaweera",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "TAPRO LAP 063",
      //           "User Name": "Kakulu Witharanage",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-03:44:00",
      //           "Install Date": "2025-10-07T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "TAPRO LAP 033",
      //           "User Name": "Pubudu Perera",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-04:13:00",
      //           "Install Date": "2025-10-07T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "TAPRO LAP 028",
      //           "User Name": "Manulya Egodapitiya",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-04:44:00",
      //           "Install Date": "2025-10-07T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "TAPRO LAP 077",
      //           "User Name": "Vimarsha Vithanage",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-04:37:00",
      //           "Install Date": "2025-10-10T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "TAPRO LAP 040 ",
      //           "User Name": "Vinuka Navod",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-03:32:00",
      //           "Install Date": "2025-10-07T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "TAPRO LAP 027",
      //           "User Name": "Rashmika Samarasinghe",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-02:10:00",
      //           "Install Date": "2025-10-07T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "TAPRO LAP 037",
      //           "User Name": "Isuru Senanayake",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "TAPRO LAP 047",
      //           "User Name": "Dinith Gunasekara",
      //           "Operating System": "Windows 11",
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": "2025-10-07T00:00:00",
      //           Notes: "Activation Key Issue",
      //         },
      //         {
      //           "Asset Tag": "TAPRO LAP 031",
      //           "User Name": "Chamika Samarawickrama",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-03:41:00",
      //           "Install Date": "2025-10-09T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "TAPRO LAP 019 ",
      //           "User Name": "Asanka Dilruk",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-04:08:00",
      //           "Install Date": "2025-10-08T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "TAPRO LAP 073",
      //           "User Name": "Anjana Gamage",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "TAPRO LAP 049",
      //           "User Name": "Ravindu Chathuranga",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-21:10:00",
      //           "Install Date": "2025-10-07T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 739 ",
      //           "User Name": "Shanuka Hettiarachchi",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-14:53:00",
      //           "Install Date": "2025-10-07T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 847",
      //           "User Name": "Hansaja Sandeepa",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-03:44:00",
      //           "Install Date": "2025-10-08T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 776",
      //           "User Name": "Thanish Ahamed",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": null,
      //           "Install Date": "2025-07-01T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 824",
      //           "User Name": "Udana Wijesuriya",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 740",
      //           "User Name": "Lasith Jayasinghe",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 783",
      //           "User Name": "Ruvindu Madushanka",
      //           "Operating System": "Windows 11",
      //           Status: null,
      //           "Installation Time": "-03:34:00",
      //           "Install Date": "2025-10-10T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 771",
      //           "User Name": "Govindhasamy Suganthini",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": null,
      //           "Install Date": "2025-07-01T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 749",
      //           "User Name": "Sewmi Punsara",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": null,
      //           "Install Date": "2025-07-01T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 802",
      //           "User Name": "Chamishka Dissanayake",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-03:13:00",
      //           "Install Date": "2025-10-08T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 819",
      //           "User Name": "Ruchira Sandaruwan",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-04:31:00",
      //           "Install Date": "2025-10-10T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 800",
      //           "User Name": "Sanka Dinesh",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "Onsite",
      //         },
      //         {
      //           "Asset Tag": "LAP 821",
      //           "User Name": "Gimhan Priyantha",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-05:38:00",
      //           "Install Date": "2025-10-10T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": null,
      //           "User Name": "Mithila Chathuranga",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 790",
      //           "User Name": "Gimhani Rubasinghe",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "Onsite",
      //         },
      //         {
      //           "Asset Tag": "LAP 513",
      //           "User Name": "Dineth Tharushan",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "Onsite",
      //         },
      //         {
      //           "Asset Tag": "LAP 660",
      //           "User Name": "Omaya Wickramasinghe",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 758",
      //           "User Name": "Manchulani Sivendrakalanithy",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "TAPRO LAP 053",
      //           "User Name": "Ovindu Nambukara",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 883",
      //           "User Name": "Nayanathara Sri Ahangama",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-02:56:00",
      //           "Install Date": "2025-10-08T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 605",
      //           "User Name": "Asanka Amarasinghe",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 571",
      //           "User Name": "Himashini Wimalasundera",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 565",
      //           "User Name": "Yasan Madhuranga",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 603",
      //           "User Name": "Madhusanka Perera",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 504",
      //           "User Name": "Yadeesha Karunathilake",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-23:25:00",
      //           "Install Date": "2025-10-09T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 604",
      //           "User Name": "Ishara Sandeepanie",
      //           "Operating System": null,
      //           Status: "Installed",
      //           "Installation Time": null,
      //           "Install Date": "2025-05-30T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 673",
      //           "User Name": "Poshitha Miguntanna",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "Onsite",
      //         },
      //         {
      //           "Asset Tag": "LAP 857",
      //           "User Name": "Upeksha Udayarathne",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-03:51:00",
      //           "Install Date": "2025-10-08T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "TAPRO LAP O32",
      //           "User Name": "Hashan lqbal",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-08:16:00",
      //           "Install Date": "2025-10-08T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "TAPRO LAP 080",
      //           "User Name": "Adeesha Thenuwara ",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": null,
      //           "Install Date": "2025-10-08T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 874",
      //           "User Name": "Supun Ramanayaka",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": null,
      //           "Install Date": "2025-10-08T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 807",
      //           "User Name": "Palinda Megasooriya",
      //           "Operating System": "Windows 10",
      //           Status: "Installed",
      //           "Installation Time": "-03:00:00",
      //           "Install Date": "2025-10-08T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "TAPRO LAP 014",
      //           "User Name": "Sayuru Disanayake",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-03:14:00",
      //           "Install Date": "2025-10-08T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 706",
      //           "User Name": "Niroshan Rathnayeka",
      //           "Operating System": "Windows 10",
      //           Status: "Installed",
      //           "Installation Time": null,
      //           "Install Date": "2025-10-08T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 869",
      //           "User Name": "Raza Shah",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-03:35:00",
      //           "Install Date": "2025-10-08T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "TAPRO LAP 064",
      //           "User Name": "Navidu Roshika",
      //           "Operating System": "Windows 10",
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "Kaspreskey Issue",
      //         },
      //         {
      //           "Asset Tag": "LAP 764",
      //           "User Name": "Jeewanthi Abeyarathne",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-20:42:00",
      //           "Install Date": "2025-10-09T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 773",
      //           "User Name": "Lahiru Karunarathne",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-25:17:00",
      //           "Install Date": "2025-10-09T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 730",
      //           "User Name": "Tharindu Gunathilaka",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-20:00:00",
      //           "Install Date": "2025-10-09T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 583",
      //           "User Name": "Nishara Kavindi",
      //           "Operating System": "Windows 10",
      //           Status: "Installed",
      //           "Installation Time": "-20:50:00",
      //           "Install Date": "2025-10-09T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 736",
      //           "User Name": "Shalani Athukorala",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-19:46:00",
      //           "Install Date": "2025-10-09T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "TAPRO LAP 074",
      //           "User Name": "Supun Chamara",
      //           "Operating System": "Windows 10",
      //           Status: "Installed",
      //           "Installation Time": "-05:08:00",
      //           "Install Date": "2025-10-09T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 679",
      //           "User Name": "Mithila De Silva",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "Kaspreskey Issue",
      //         },
      //         {
      //           "Asset Tag": "LAP 656",
      //           "User Name": "Pansilu Nilaweera",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "Onsite",
      //         },
      //         {
      //           "Asset Tag": "LAP 655",
      //           "User Name": "Chamika Perera",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-05:47:00",
      //           "Install Date": "2025-10-09T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 697",
      //           "User Name": "Mathisha Wanigasinghe",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "Onsite",
      //         },
      //         {
      //           "Asset Tag": "LAP 703",
      //           "User Name": "Pradeep Kumara",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 715",
      //           "User Name": "Nalaka Sandaruwan",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 719",
      //           "User Name": "Charitha Rathnayaka",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 718",
      //           "User Name": "Kalana Hewabatuwita",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 723",
      //           "User Name": "Dinesh Mayura Bandara",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "TAPRO LAP 069",
      //           "User Name": "Udana Kodikara",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-08:54:00",
      //           "Install Date": "2025-10-09T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 845",
      //           "User Name": "Lakshan Chathuranga",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-02:48:00",
      //           "Install Date": "2025-10-09T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 484",
      //           "User Name": "Udara dharmadasa",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-20:13:00",
      //           "Install Date": "2025-10-09T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "TAPRO LAP 062",
      //           "User Name": "Tharushi Hindakaraldeniya",
      //           "Operating System": "Windows 10",
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "Kaspreskey Issue",
      //         },
      //         {
      //           "Asset Tag": "TAPRO LAP 071",
      //           "User Name": "Sharanga De Alwis",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-05:14:00",
      //           "Install Date": "2025-10-09T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 785",
      //           "User Name": "Nimmi Senanayake",
      //           "Operating System": "Windows 10",
      //           Status: "Installed",
      //           "Installation Time": "-04:53:00",
      //           "Install Date": "2025-10-09T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "TAPRO LAP 018",
      //           "User Name": "Isuru Dissanayake",
      //           "Operating System": "Windows 11",
      //           Status: "Installed",
      //           "Installation Time": "-05:21:00",
      //           "Install Date": "2025-10-10T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "TAPRO LAP 065",
      //           "User Name": "Damith Atharagalla",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "TAPRO LAP 015",
      //           "User Name": "Jeewaka Manohara",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 611",
      //           "User Name": "Mohamed Ramsan",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 868",
      //           "User Name": "Chamath Amarasuriya",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 756",
      //           "User Name": "Achintha Dissanayaka",
      //           "Operating System": "Windows 10",
      //           Status: "Installed",
      //           "Installation Time": "-05:13:00",
      //           "Install Date": "2025-10-10T00:00:00",
      //           Notes: "Onsite",
      //         },
      //         {
      //           "Asset Tag": "LAP 766",
      //           "User Name": "Jayathi Yasara Ransisi",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 759",
      //           "User Name": "Sajith Madhusankha",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 784",
      //           "User Name": "Ruwan Siriwardena",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 777",
      //           "User Name": "Lakshitha Perera",
      //           "Operating System": "Windows 10",
      //           Status: "Installed",
      //           "Installation Time": null,
      //           "Install Date": "2025-10-10T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 789",
      //           "User Name": "Tinali Gunasekara",
      //           "Operating System": "Windows 10",
      //           Status: "Installed",
      //           "Installation Time": "-04:47:00",
      //           "Install Date": "2025-10-10T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 793",
      //           "User Name": "Sasindu Prasad",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 797",
      //           "User Name": "Thisara Pannala",
      //           "Operating System": "Windows 10",
      //           Status: "Installed",
      //           "Installation Time": "-03:51:00",
      //           "Install Date": "2025-10-10T00:00:00",
      //           Notes: "Done",
      //         },
      //         {
      //           "Asset Tag": "LAP 801",
      //           "User Name": "Deshan Balasuriya",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 809",
      //           "User Name": "Sahan Kariyawasam",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 814",
      //           "User Name": "Udayanga Amarathunga",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 820",
      //           "User Name": "Binuka Kamesh",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 826",
      //           "User Name": "Sandun Pathirana",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 827",
      //           "User Name": "Kaveenda Navodyanga",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 828",
      //           "User Name": "Ravindu Maginaarachchi",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 870",
      //           "User Name": "Lakshan Mihiranga",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 695",
      //           "User Name": "Shehani Silva",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 837",
      //           "User Name": "Hashini Kaushalya",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 843",
      //           "User Name": "Kasun Gamage",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 844",
      //           "User Name": "Ishani Perera",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "LAP 816",
      //           "User Name": "Madushan Sandaruwan",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "TAPRO LAP 038",
      //           "User Name": "Keshan Sankalpa",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "TAPRO LAP 059",
      //           "User Name": "Hasantha Pathirana",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "TAPRO LAP 014",
      //           "User Name": "Sayuru Dissanayake",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "TAPRO LAP 044",
      //           "User Name": "Darshana Fernando",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //         {
      //           "Asset Tag": "TAPRO LAP 012",
      //           "User Name": "Kumudu Weerasinghe",
      //           "Operating System": null,
      //           Status: null,
      //           "Installation Time": null,
      //           "Install Date": null,
      //           Notes: "No Response",
      //         },
      //       ],
      //     },
      //     Dashboard_Metrics: {
      //       row_count: 16,
      //       column_count: 2,
      //       columns: ["Metric", "Value"],
      //       data: [
      //         {
      //           Metric: "Total Assets",
      //           Value: "220",
      //         },
      //         {
      //           Metric: "Installed",
      //           Value: "146",
      //         },
      //         {
      //           Metric: "No Response",
      //           Value: "61",
      //         },
      //         {
      //           Metric: "Issues",
      //           Value: "4",
      //         },
      //         {
      //           Metric: "Completed",
      //           Value: "146",
      //         },
      //         {
      //           Metric: "Protected",
      //           Value: "146",
      //         },
      //         {
      //           Metric: "Remaining",
      //           Value: "74",
      //         },
      //         {
      //           Metric: "Start Date",
      //           Value: "2025-10-02T00:00:00",
      //         },
      //         {
      //           Metric: "Days Active",
      //           Value: "11",
      //         },
      //         {
      //           Metric: "Last Install",
      //           Value: "2025-10-10T00:00:00",
      //         },
      //         {
      //           Metric: "Success Rate",
      //           Percentage: "66",
      //         },
      //         {
      //           Metric: "Response Rate",
      //           Percentage: "72",
      //         },
      //         {
      //           Metric: "Issue Rate",
      //           Percentage: "2",
      //         },
      //         {
      //           Metric: "Pending Rate",
      //           Percentage: "34",
      //         },
      //         {
      //           Metric: "Completion Rate",
      //           Percentage: "66",
      //         },
      //         {
      //           Metric: "Avg Per Day",
      //           Value: "27",
      //         },
      //       ],
      //     },
      //     OS_Distribution: {
      //       row_count: 3,
      //       column_count: 3,
      //       columns: ["Operating System", "Count", "Percentage"],
      //       data: [
      //         {
      //           "Operating System": "Windows 11",
      //           Count: "126",
      //           Percentage: "85",
      //         },
      //         {
      //           "Operating System": "Windows 10",
      //           Count: "20",
      //           Percentage: "13",
      //         },
      //         {
      //           "Operating System": "macOS",
      //           Count: "3",
      //           Percentage: "2",
      //         },
      //       ],
      //     },
      //     Status_Breakdown: {
      //       row_count: 5,
      //       column_count: 4,
      //       columns: ["Status Category", "Count", "Percentage", "Notes"],
      //       data: [
      //         {
      //           "Status Category": "Installed",
      //           Count: "146",
      //           Percentage: "66",
      //           Notes: "Successfully deployed",
      //         },
      //         {
      //           "Status Category": "Completed",
      //           Count: "146",
      //           Percentage: "66",
      //           Notes: "Fully operational",
      //         },
      //         {
      //           "Status Category": "No Response",
      //           Count: "61",
      //           Percentage: "28",
      //           Notes: "Awaiting user action",
      //         },
      //         {
      //           "Status Category": "Issues",
      //           Count: "4",
      //           Percentage: "2",
      //           Notes: "Requires attention",
      //         },
      //         {
      //           "Status Category": "Remaining",
      //           Count: "74",
      //           Percentage: "34",
      //           Notes: "Not yet deployed",
      //         },
      //       ],
      //     },
      //     Risk_Assessment: {
      //       row_count: 3,
      //       column_count: 3,
      //       columns: ["Risk Level", "Status", "Description"],
      //       data: [
      //         {
      //           "Risk Level": "High Risk",
      //           Status: null,
      //           Description: "No critical risks identified",
      //         },
      //         {
      //           "Risk Level": "Medium Risk",
      //           Status: "Active",
      //           Description:
      //             "Delayed completion may impact compliance requirements and organizational security posture. Current deployment leaves significant coverage gaps.",
      //         },
      //         {
      //           "Risk Level": "Mitigation Strategy",
      //           Status: "In Progress",
      //           Description:
      //             "Enhanced network monitoring activated for unprotected devices. Priority escalation for manual installations scheduled within 7 days to minimize exposure window.",
      //         },
      //       ],
      //     },
      //   },
      // };

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
    if (!this.excelData || !this.excelData.sheets) {
      return this.getEmptyDashboardFormat();
    }

    const sheets = this.excelData.sheets;

    // Get metrics from Dashboard_Metrics sheet
    const metricsData = sheets.Dashboard_Metrics?.data || [];
    const metrics = {};
    metricsData.forEach((item) => {
      const key = item.Metric?.toLowerCase().replace(/\s+/g, "");
      metrics[key] = item.Value || item.Percentage || 0;
    });

    // Get OS distribution
    const osData = sheets.OS_Distribution?.data || [];
    const osDistribution = osData.map((item) => ({
      name: item["Operating System"] || "Unknown",
      value: parseInt(item.Count) || 0,
      percentage: parseFloat(item.Percentage || 0),
      version: item["Operating System"]?.includes("11") ? "22H2" : "22H2",
      security: item["Operating System"] === "macOS" ? "High" : "High",
    }));

    // Get status breakdown
    const statusData = sheets.Status_Breakdown?.data || [];
    const statusBreakdown = {};
    statusData.forEach((item) => {
      const key =
        item["Status Category"]?.toLowerCase().replace(/\s+/g, "") || "unknown";
      statusBreakdown[key] = parseInt(item.Count) || 0;
    });

    // Get device inventory - CRITICAL SECTION
    const deviceData = sheets.Device_Inventory?.data || [];
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
    const riskData = sheets.Risk_Assessment?.data || [];
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
    const deviceDataMain =
      this.dashboardData?.rawWorksheets?.Device_Inventory?.data || [];

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
    this.pageInfo.textContent = `Page ${this.currentPage} of ${Math.ceil(this.totalDevicesCount / this.recordsPerPage)}`;

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
            
              <span class="os-name">${device["Operating System"] || "Unknown"}</span>
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
        this.renderOSChart(),
        this.updateInsights(),
      ]);
    } catch (error) {
      console.error("Error updating dashboard:", error);
    }
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
paginatorBtn.addEventListener("change", e => {
  window.dashboard.setPaginatorElements(1,e.target.value)
  window.dashboard.populateDevicesTable();
})
