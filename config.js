const CONFIG = {

  // You will fill these in during the Azure setup step
  clientId:    'PASTE_YOUR_CLIENT_ID_HERE',
  tenantId:    'PASTE_YOUR_TENANT_ID_HERE',
  redirectUri: 'https://JadeABA.github.io/jade-aba-ceo-dashboard',

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
      cells: {
        totalAuthorized:   'B4',
        totalAvailable:    'B5',
        totalScheduled:    'B6',
        overallUtilPct:    'B7',
        clientsUnder80:    'J5',
        therapistsUnder70: 'J6',
        openHours:         'J7',
      }
    }

  },

  companyName:          'Jade ABA Therapy',
  refreshMinutes:       30,
  utilizationTarget:    85,
  collectionTarget:     88,
  bcbaTarget:           100,
  contactSpeedTarget:   24,
};
