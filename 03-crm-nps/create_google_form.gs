/**
 * INDIAN NEWS & BUSINESS MAGAZINE READER SURVEY — FORM BUILDER
 * ---------------------------------------------------------------
 * HOW TO USE:
 * 1. Go to https://script.google.com -> New project.
 * 2. Delete any starter code in Code.gs, paste this whole file in.
 * 3. Click the "Run" button (▷) with function createSurveyForm selected.
 * 4. First run: Google will ask you to authorize -> click through
 *    "Advanced" -> "Go to (project name) (unsafe)" -> Allow.
 *    (This warning is normal for your own scripts; you wrote it.)
 * 5. Check the "Execution log" (View > Logs) for two URLs:
 *    - Editor URL (for you, to review/tweak questions)
 *    - Live form URL (this is what you distribute)
 * 6. Responses land automatically in a new Google Sheet named
 *    "Magazine Reader Survey - Responses", linked to the form.
 */

function createSurveyForm() {
  const form = FormApp.create('Indian News & Business Magazine Reader Survey')
    .setDescription(
      'This short survey (4-5 minutes) is part of an MBA business analytics research ' +
      'project on reader loyalty and subscription habits for Indian news and business ' +
      'magazines. Your responses are anonymous and used for academic analysis only. ' +
      'Thank you for participating!'
    )
    .setCollectEmail(false)
    .setAllowResponseEdits(false)
    .setLimitOneResponsePerUser(false)
    .setProgressBar(true)
    .setShuffleQuestions(false);

  // ---------- SECTION 1: Screening & Demographics ----------
  form.addSectionHeaderItem()
    .setTitle('Section 1: A Few Quick Details');

  form.addMultipleChoiceItem()
    .setTitle('Do you currently read, or have you read in the past, any Indian news or ' +
      'business magazine (print or digital)?')
    .setChoiceValues(['Yes', 'No'])
    .setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle('What is your age bracket?')
    .setChoiceValues(['Under 18', '18-25', '26-35', '36-50', '51-65', '65+'])
    .setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle('What is your gender?')
    .setChoiceValues(['Male', 'Female', 'Prefer not to say', 'Other'])
    .setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle('Which part of India do you live in?')
    .setChoiceValues(['North India', 'South India', 'East India', 'West India',
      'Central India', 'Outside India'])
    .setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle('Which best describes you?')
    .setChoiceValues(['Student', 'Salaried professional', 'Business owner or self-employed',
      'Freelancer', 'Retired', 'Other'])
    .setRequired(true);

  // ---------- SECTION 2: Customer Acquisition ----------
  form.addPageBreakItem()
    .setTitle('Section 2: How You Started Reading');

  form.addCheckboxItem()
    .setTitle('Which Indian news/business magazine(s) do you currently read or subscribe to? ' +
      '(select all that apply)')
    .setChoiceValues(['India Today', 'Outlook', 'Business Today', 'The Week',
      'Forbes India', 'Frontline', 'Open Magazine', 'Other'])
    .setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle('How did you first discover this magazine?')
    .setChoiceValues(['Family or friend recommendation', 'Social media', 'Online search',
      'Newsstand or vendor', 'Promotional offer or discount',
      'Bundled with another subscription', 'Other'])
    .setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle('What was the primary reason you first started reading or subscribed?')
    .setChoiceValues(['Content quality', 'Price or promotional offer', 'Habit or brand legacy',
      'Convenience of access', 'Influenced by family, friends, or colleagues', 'Other'])
    .setRequired(true);

  // ---------- SECTION 3: Customer Retention ----------
  form.addPageBreakItem()
    .setTitle('Section 3: Your Ongoing Reading Habits');

  form.addMultipleChoiceItem()
    .setTitle('How long have you been reading or subscribed to this magazine?')
    .setChoiceValues(['Less than 6 months', '6 months to 1 year', '1-3 years',
      '3-5 years', 'More than 5 years'])
    .setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle('How do you currently access it?')
    .setChoiceValues(['Print only', 'Digital only', 'Both print and digital'])
    .setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle('How thoroughly do you typically read an issue?')
    .setChoiceValues(['Read it cover-to-cover', 'Skim most articles',
      'Only read specific sections'])
    .setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle('Have you ever seriously considered cancelling or not renewing a subscription?')
    .setChoiceValues(['Yes', 'No'])
    .setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle('If yes above - what was the main reason? (Choose "Not applicable" if you ' +
      'answered No)')
    .setChoiceValues(['Price increase', 'Decline in content quality', 'Lack of time to read',
      'Switched to another source', 'Not applicable'])
    .setRequired(true);

  // ---------- SECTION 4: Customer Development ----------
  form.addPageBreakItem()
    .setTitle('Section 4: Bundles & Upgrades');

  form.addMultipleChoiceItem()
    .setTitle('Have you ever subscribed to more than one magazine from the same publisher?')
    .setChoiceValues(['Yes', 'No'])
    .setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle('Would you be interested in a bundled offer (e.g. print + digital + a second ' +
      'title at a discount)?')
    .setChoiceValues(['Yes', 'No', 'Maybe, depends on price'])
    .setRequired(true);

  form.addCheckboxItem()
    .setTitle('Which promotional offers would most likely make you upgrade or renew? ' +
      '(select all that apply)')
    .setChoiceValues(['Extra free months on renewal', 'Discounted multi-title bundle',
      'Free gift or merchandise', 'Referral rewards', 'Exclusive extra digital content',
      'A price-lock guarantee against future increases'])
    .setRequired(true);

  // ---------- SECTION 5: Net Promoter Score ----------
  form.addPageBreakItem()
    .setTitle('Section 5: One Last Set of Questions');

  form.addScaleItem()
    .setTitle('On a scale of 0-10, how likely are you to recommend this magazine to a ' +
      'friend or colleague?')
    .setBounds(0, 10)
    .setLabels('Not at all likely', 'Extremely likely')
    .setRequired(true);

  form.addParagraphTextItem()
    .setTitle('What is the main reason for the score you gave above?')
    .setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle('What one thing would most improve your experience with this magazine?')
    .setChoiceValues(['Lower price', 'Greater content depth', 'Faster digital delivery',
      'More regional/local coverage', 'Better customer service', 'Other'])
    .setRequired(true);

  // ---------- SECTION 6: International Comparison ----------
  form.addPageBreakItem()
    .setTitle('Section 6: Almost Done');

  form.addMultipleChoiceItem()
    .setTitle('Have you ever subscribed to an international publication (e.g. The Economist, ' +
      'TIME, The New York Times, Financial Times)?')
    .setChoiceValues(['Yes, currently', 'Yes, in the past', 'No'])
    .setRequired(true);

  form.addTextItem()
    .setTitle('If yes - which international publication(s)? (leave blank if not applicable)')
    .setRequired(false);

  // ---------- Link responses to a new Google Sheet ----------
  const ss = SpreadsheetApp.create('Magazine Reader Survey - Responses');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  Logger.log('EDITOR URL (for you): ' + form.getEditUrl());
  Logger.log('LIVE FORM URL (to distribute): ' + form.getPublishedUrl());
  Logger.log('RESPONSES SHEET URL: ' + ss.getUrl());
}
