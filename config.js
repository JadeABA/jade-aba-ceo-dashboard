const CONFIG = {

  clientId:    '5a8e24ef-86dc-443d-93fc-e74b11365575',
  tenantId:    '33ee9db4-64f2-4c9f-a3c3-ae3cd287d483',
  redirectUri: 'https://JadeABA.github.io/jade-aba-ceo-dashboard',

  files: {

    ar: {
      fileId:    '01SYCMTOTGJJPJ5SWVBFAJ3KIW6XPT3CRH',
      sheetName: 'Weekly AR Tracker',
      columns: {
        weekStart:        'Week Start Date',
        weekEnd:          'Week End Date',
        ar0_30:           'AR 0-30',
        ar31_60:          'AR 31-60',
        ar61_90:          'AR 61-90',
        ar90plus:         'AR 90+',
        totalAR:          'Total AR',
        claimsSubmitted:  'Claims Submitted',
        denialsReceived:  'Denials Received',
        denialsPending:   'Denials Pending',
        cleanClaimRate:   'Clean Claim Rate',
        avgDaysPayment:   'Avg Days to Payment',
        totalWeeklyPay:   'Total Weekly Payment',
      }
    },

    bcba: {
      fileId:    '01SYCMTOQC6HFL64F7FNC2QGKLFTXYAZOZ',
      sheetName: 'BCBA Weekly Tracker',
      columns: {
        weekStart:     'Week Start Date',
        bcbaName:      'BCBA Name',
        billableHours: 'Billable Hours',
        target:        'Target (30)',
        pctToTarget:   '% to Target',
        activeCases:   'Active Cases',
      }
    },

    intake: {
      fileId:    '01SYCMTOSMYF3XPYHBUFCJONPOS6V5V2Z5',
      sheetName: 'Intake',
      columns: {
        weekStart:        'Week Start',
        weekEnd:          'Week End',
        newLeads:         'New Leads',
        contacted24h:     'Contacted <24h',
        pctContacted:     '% Contacted',
        intakesCompleted: 'Intakes Completed',
        leadsInPool:      'Total Leads in Pool',
        conversionPct:    'Conversion %',
        avgDaysToIntake:  'Avg Days to Intake',
        pendingDocs:      'Pending Documents',
        referralOutreach: 'Referral Outreach',
        referralFollowup: 'Referral Followups',
      }
    },

    recruiting: {
      fileId:    '01SYCMTOSMYF3XPYHBUFCJONPOS6V5V2Z5',
      sheetName: 'Recruiting',
      columns: {
        weekStart:    'Week Start',
        weekEnd:      'Week End',
        openBCBA:     'Open BCBA',
        openRBT:      'Open RBT',
        applications: 'Applications',
        interviews:   'Interviews',
        offers:       'Offers',
        hires:        'Hires',
        hireRate:     'Hire Rate',
        timeToFill:   'Time to Fill RBT (days)',
        casesDelayed: 'Cases Delayed',
      }
    },

    caseCoord: {
      fileId:    '01SYCMTOSMYF3XPYHBUFCJONPOS6V5V2Z5',
      sheetName: 'Case Coordination',
      columns: {
        weekStart:      'Week Start',
        weekEnd:        'Week End',
        activeCases:    'Active Cases',
        unstaffedCases: 'Unstaffed Cases',
        casesLaunched:  'Cases Launched',
        avgDaysIntake:  'Avg Days Intake\u2192Assessment',
      }
    },

    utilization: {
      fileId:    '01SYCMTOTKJVIYNS55UVBLBAH2GUICBXBO',
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

  companyName:         'Jade ABA Therapy',
  refreshMinutes:      30,
  utilizationTarget:   85,
  collectionTarget:    88,
  bcbaTarget:          100,
  contactSpeedTarget:  24,
};
