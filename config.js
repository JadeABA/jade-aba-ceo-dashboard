const CONFIG = {

  // You will fill these in during Step 3 (Azure setup)
  clientId:    'PASTE_YOUR_CLIENT_ID_HERE',
  tenantId:    'PASTE_YOUR_TENANT_ID_HERE',
  redirectUri: 'https://YOUR-GITHUB-USERNAME.github.io/jade-aba-ceo-dashboard',

  // You will fill these in during Step 4 (getting File IDs from OneDrive)
  files: {

    ar: {
      fileId:    'PASTE_AR_FILE_ID_HERE',
      sheetName: 'Weekly AR Tracker',
      columns: {
        weekStart:   'Week Start Date',
        weekEnd:     'Week End Date',
        ar0_30:      'AR 0-30',
        ar31_60:     'AR 31-60',
        ar61_90:     'AR 61-90',
        ar90plus:    'AR 90+',
        totalAR:     'Total AR',
      }
    },

    bcba: {
      fileId:    'PASTE_KPI_FILE_ID_HERE',
      sheetName: 'BCBA Weekly Tracker',
      columns: {
        weekStart:     'Week Start Date',
        bcbaName:      'BCBA Name',
        billableHours: 'Billable Hours',
        target:        'Target (30)',
        pctToTarget:   '% to Target',
      }
    },

    intake: {
      fileId:    'PASTE_OPERATIONS_FILE_ID_HERE',
      sheetName: 'Intake',
      columns: {
        weekStart:        'Week Start',
        weekEnd:          'Week End',
        newLeads:         'New Leads',
        contacted24h:     'Contacted <24h',
        pctContacted:     '% Contacted',
        intakesCompleted: 'Intakes Completed',
        conversionPct:    'Conversion %',
        avgDaysToIntake:  'Avg Days to Intake',
        pendingDocs:      'Pending Documents',
        referralOutreach: 'Referral Outreach',
        referralFollowup: 'Referral Followups',
      }
    },

    recruiting: {
      fileId:    'PASTE_OPERATIONS_FILE_ID_HERE',
      sheetName: 'Recruiting',
      columns: {
        weekStart:       'Week Start',
        weekEnd:         'Week End',
        openBCBA:        'Open BCBA',
        openRBT:         'Open RBT',
        applications:    'Applications',
        interviews:      'Interviews',
        offers:          'Offers',
        hires:           'Hires',
        hireRate:        'Hire Rate',
        timeToFill:      'Time to Fill RBT (days)',
        casesDelayed:    'Cases Delayed',
      }
    },

    caseCoord: {
      fileId:    'PASTE_OPERATIONS_FILE_ID_HERE',
      sheetName: 'Case Coordination',
      columns: {
        weekStart:       'Week Start',
        weekEnd:         'Week End',
        activeCases:     'Active Cases',
        unstaffedCases:  'Unstaffed Cases',
        casesLaunched:   'Cases Launched',
        avgDaysIntake:   'Avg Days Intake→Assessment',
      }
    },

    utilization: {
      fileId:    'PASTE_UTILIZATION_FILE_ID_HERE',
      sheetName: 'Weekly_Utilization_Dashboard',
      // These are the exact cell positions in your summary dashboard tab
      cells: {
        totalAuthorized:  'B4',
        totalAvailable:   'B5',
        totalScheduled:   'B6',
        overallUtilPct:   'B7',
        clientsUnder80:   'J5',
        therapistsUnder70:'J6',
        openHours:        'J7',
        clientTableStart:  12,   // Row where client list begins
        therapistColStart: 'F',  // Column where therapist table starts
      }
    }

  },

  // Dashboard targets — green/yellow/red thresholds
  companyName:          'Jade ABA Therapy',
  refreshMinutes:       30,
  utilizationTarget:    85,       // % — green at or above this
  collectionTarget:     88,       // % collection rate target
  bcbaTarget:           100,      // % to target — green at or above
  contactSpeedTarget:   24,       // hours to contact new leads
};
