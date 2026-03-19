/**
 * 谷雨暑期选修课报名表 — Google Apps Script
 * 自动创建表单 + 自动计算学费 + 邮件确认
 *
 * 使用方法 How to use:
 * 1. 打开 https://script.google.com
 * 2. 新建项目 → 删除默认代码 → 粘贴此脚本
 * 3. 点击 ▶ Run → 选择 createRegistrationForm → 授权
 * 4. 查看日志 (View → Execution log) 获取表单链接
 * 5. 表单创建后，脚本会自动设置 onFormSubmit 触发器
 *    每次家长提交表单后，系统自动计算学费并发送确认邮件
 */

function createRegistrationForm() {
  var form = FormApp.create('谷雨暑期选修课报名 GR EDU Summer Electives Registration');
  form.setDescription(
    '2025 Summer · Chinese · English · Math\n' +
    '请填写以下信息完成报名。Please fill out the form below to register.\n' +
    '如需报名多个学生，请为每位学生分别提交一份表格。Submit one form per student.'
  );
  form.setCollectEmail(true);
  form.setConfirmationMessage(
    '报名表已提交！Thank you for registering!\n' +
    '系统将自动计算学费并发送确认邮件到您的邮箱。\n' +
    'A confirmation email with tuition details will be sent to your email shortly.\n\n' +
    '请尽快完成 Zelle 付款，名额以付款先后顺序确认。'
  );

  // ====== SECTION 1: Parent & Student Info ======
  form.addSectionHeaderItem()
    .setTitle('1. 家长与学生信息 Parent & Student Info');

  form.addTextItem()
    .setTitle('家长姓名 Parent Name')
    .setRequired(true);

  form.addTextItem()
    .setTitle('联系电话 Phone')
    .setRequired(true);

  form.addTextItem()
    .setTitle('学生姓名 Student Name')
    .setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle('学生年级 Student Grade (2026 Fall 入学年级)')
    .setChoiceValues([
      'Kindergarten (K)',
      '1st Grade',
      '2nd Grade',
      '3rd Grade',
      '4th Grade',
      '5th Grade'
    ])
    .setRequired(true);

  // ====== SECTION 2: Select Site ======
  form.addPageBreakItem()
    .setTitle('2. 选择校区 Select Site');

  form.addMultipleChoiceItem()
    .setTitle('上课地点 Campus')
    .setChoiceValues(['Cupertino', 'Milpitas'])
    .setRequired(true);

  // ====== SECTION 3: Select Sessions ======
  form.addPageBreakItem()
    .setTitle('3. 选择 Session')
    .setHelpText(
      '可多选 Select all that apply\n' +
      '每个 Session 共 4 节课（每周 2 次 × 2 周）。选修课在所有 Session 的时间和级别相同。\n' +
      'Each session = 4 classes. Schedule and levels are the same across all sessions.'
    );

  form.addCheckboxItem()
    .setTitle('报名 Session')
    .setChoiceValues([
      'Session 1 (Week 1: 6/8–6/12, Week 2: 6/15–6/19)',
      'Session 2 (Week 3: 6/22–6/26, Week 4: 7/6–7/10)',
      'Session 3 (Week 5: 7/13–7/17, Week 6: 7/20–7/24)',
      'Session 4 (Week 7: 7/27–7/31, Week 8: 8/3–8/7)'
    ])
    .setRequired(true);

  // ====== SECTION 4: Select Courses & Levels ======
  form.addPageBreakItem()
    .setTitle('4. 选择课程与级别 Select Courses & Levels')
    .setHelpText('每科最多选一个级别 One level per subject · 所有时间均为 PM (PST) · Onsite Classes');

  // Chinese
  form.addMultipleChoiceItem()
    .setTitle('中文级别 Chinese Level')
    .setHelpText('中文课按实际中文水平分级，非按学校年级分班。L1 无需测评，L2–L5 建议完成入学测评后选择。')
    .setChoiceValues([
      '不选中文 N/A',
      'Chinese L1（零基础）— Tue/Thu 5:00–6:00 — 参考年级: K–1st — 无需测评',
      'Chinese L2 — Mon/Wed 4:00–5:00 — 参考年级: 1st–2nd',
      'Chinese L3 — Mon/Wed 5:00–6:00 — 参考年级: 2nd–3rd',
      'Chinese L4 — Fri 4:00–6:00 — 参考年级: 3rd–5th',
      'Chinese L5 — Tue/Thu 4:00–5:00 — 参考年级: 4th+'
    ])
    .setRequired(true);

  // English
  form.addMultipleChoiceItem()
    .setTitle('英语级别 English Level')
    .setChoiceValues([
      '不选英语 N/A',
      'English LK — Mon/Wed 4:00–5:00 — Grade: K',
      'English L1 — Mon/Wed 5:00–6:00 — Grade: 1st',
      'English L2 & 3 — Tue/Thu 4:00–5:00 — Grade: 2nd–3rd',
      'English L4 & 5 — Tue/Thu 5:00–6:00 — Grade: 3rd–5th'
    ])
    .setRequired(true);

  // Math
  form.addMultipleChoiceItem()
    .setTitle('数学级别 Math Level')
    .setHelpText('数学课按实际水平分级，非按学校年级分班。建议查看各级别课程内容后选择。')
    .setChoiceValues([
      '不选数学 N/A',
      'Math L1 — Tue/Thu 4:00–5:00 — 参考年级: 1st',
      'Math L2 — Tue/Thu 5:00–6:00 — 参考年级: 2nd–3rd',
      'Math L3 — Mon/Wed 4:00–5:00 — 参考年级: 3rd',
      'Math L4 — Mon/Wed 5:00–6:00 — 参考年级: 4th–5th',
      'Math L5 — Fri 4:00–6:00 — 参考年级: 4th–5th'
    ])
    .setRequired(true);

  // ====== SECTION 5: Tuition Reference ======
  form.addPageBreakItem()
    .setTitle('5. 学费参考 Tuition Reference (系统将自动计算)')
    .setHelpText(
      '学费标准 Tuition per session (2 weeks):\n' +
      '• 中文 Chinese: $120/session ($60/week)\n' +
      '• 英语 English: $160/session ($80/week)\n' +
      '• 数学 Math: $160/session ($80/week)\n\n' +
      '学费 = 所选科目单价 × 报名 Session 数\n' +
      'Total = per-session price × number of sessions\n\n' +
      '提交表单后，系统将自动计算学费并发送确认邮件。\n' +
      'Tuition will be auto-calculated and emailed to you after submission.'
    );

  // ====== SECTION 6: Refund Policy ======
  form.addPageBreakItem()
    .setTitle('6. 退费政策 Refund Policy')
    .setHelpText(
      '请仔细阅读后再付款 Please read before payment\n\n' +
      '退费说明 Refund Policy:\n' +
      '• 2026年4月15日之前，已支付费用可 100% 转为学校 credit，用于报名谷雨其他课程（夏令营、课后班、周末课等）。\n' +
      '  Before April 15, 2026: full credit towards any GR EDU program.\n' +
      '• 2026年4月15日之后，选修课为小班教学，涉及分班、教师安排及开班协调，如因个人原因无法参加，费用将不予退还，敬请理解。\n' +
      '  After April 15, 2026: no refunds due to class planning and staffing commitments.\n\n' +
      '名额转让 Transfer Policy:\n' +
      '• 开营前，可按 Session 将名额转让给他人，不收取 processing fee。\n' +
      '  Before camp starts: transfer your spot by session to another student, no fee.'
    );

  form.addMultipleChoiceItem()
    .setTitle('我已阅读并同意以上退费政策 I have read and agree to the refund policy')
    .setChoiceValues(['是的，我已阅读并同意 Yes, I agree'])
    .setRequired(true);

  // ====== SECTION 7: Payment ======
  form.addPageBreakItem()
    .setTitle('7. 缴费方式 Payment')
    .setHelpText(
      'Zelle 转账 Payment via Zelle:\n' +
      '收款账号 Send to: gredu2019@gmail.com\n' +
      '备注 Memo: 学生姓名 + 选修课 Elective\n' +
      '（例: John Wang - Summer Elective）\n\n' +
      '提交报名表后请尽快完成 Zelle 付款，名额以付款先后顺序确认。\n' +
      'Spots are confirmed on a first-paid basis.'
    );

  form.addMultipleChoiceItem()
    .setTitle('付款确认 Payment Confirmation')
    .setChoiceValues([
      '已通过 Zelle 付款 Paid via Zelle',
      '稍后付款（提交表格后 48 小时内完成付款）Will pay within 48 hours'
    ])
    .setRequired(true);

  form.addTextItem()
    .setTitle('Zelle 付款人姓名 Zelle Sender Name');

  form.addParagraphTextItem()
    .setTitle('备注 Additional Notes');

  // ====== Link to Google Sheet ======
  var ss = SpreadsheetApp.create('谷雨暑期选修课报名 Responses');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  // Add "应付总额 Total Due" column header to the sheet
  // (will be filled by the onFormSubmit trigger)
  var sheet = ss.getSheets()[0];
  // The form auto-creates headers; we add our calculated column after the last one
  // We'll do this in the trigger since form destination needs time to sync

  // ====== Set up auto-submit trigger ======
  ScriptApp.newTrigger('onFormSubmit')
    .forForm(form)
    .onFormSubmit()
    .create();

  Logger.log('========================================');
  Logger.log('Form created successfully!');
  Logger.log('Edit URL: ' + form.getEditUrl());
  Logger.log('Share URL: ' + form.getPublishedUrl());
  Logger.log('Sheet URL: ' + ss.getUrl());
  Logger.log('========================================');
  Logger.log('Auto-calculation trigger has been set up.');
  Logger.log('When a parent submits the form:');
  Logger.log('  1. Tuition is auto-calculated');
  Logger.log('  2. Total is written to the response sheet');
  Logger.log('  3. Confirmation email is sent to the parent');
}

/**
 * Triggered automatically when a form response is submitted.
 * Calculates tuition and sends confirmation email.
 */
function onFormSubmit(e) {
  var response = e.response;
  var items = response.getItemResponses();

  // Parse responses
  var parentName = '';
  var studentName = '';
  var phone = '';
  var grade = '';
  var campus = '';
  var sessionCount = 0;
  var chineseLevel = '';
  var englishLevel = '';
  var mathLevel = '';
  var email = response.getRespondentEmail();

  for (var i = 0; i < items.length; i++) {
    var title = items[i].getItem().getTitle();
    var answer = items[i].getResponse();

    if (title.indexOf('家长姓名') >= 0) parentName = answer;
    if (title.indexOf('学生姓名') >= 0) studentName = answer;
    if (title.indexOf('联系电话') >= 0) phone = answer;
    if (title.indexOf('学生年级') >= 0) grade = answer;
    if (title.indexOf('上课地点') >= 0) campus = answer;
    if (title.indexOf('报名 Session') >= 0) {
      // answer is an array for checkboxes
      sessionCount = Array.isArray(answer) ? answer.length : 1;
    }
    if (title.indexOf('中文级别') >= 0) chineseLevel = answer;
    if (title.indexOf('英语级别') >= 0) englishLevel = answer;
    if (title.indexOf('数学级别') >= 0) mathLevel = answer;
  }

  // Calculate tuition
  var chinesePerSession = 120;  // $60/week × 2 weeks
  var englishPerSession = 160;  // $80/week × 2 weeks
  var mathPerSession = 160;     // $80/week × 2 weeks

  var hasChinese = chineseLevel && chineseLevel.indexOf('N/A') < 0;
  var hasEnglish = englishLevel && englishLevel.indexOf('N/A') < 0;
  var hasMath = mathLevel && mathLevel.indexOf('N/A') < 0;

  var perSessionTotal = 0;
  var breakdown = [];

  if (hasChinese) {
    perSessionTotal += chinesePerSession;
    breakdown.push('中文 Chinese: $' + chinesePerSession + '/session');
  }
  if (hasEnglish) {
    perSessionTotal += englishPerSession;
    breakdown.push('英语 English: $' + englishPerSession + '/session');
  }
  if (hasMath) {
    perSessionTotal += mathPerSession;
    breakdown.push('数学 Math: $' + mathPerSession + '/session');
  }

  var totalDue = perSessionTotal * sessionCount;

  // Write total to the linked spreadsheet
  try {
    var form = FormApp.getActiveForm() || FormApp.openById(e.source.getId());
    var destId = form.getDestinationId();
    var ss = SpreadsheetApp.openById(destId);
    var sheet = ss.getSheets()[0];
    var lastCol = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();

    // Check if "应付总额 Total Due" column exists
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var totalColIndex = -1;
    for (var h = 0; h < headers.length; h++) {
      if (headers[h] === '应付总额 Total Due') {
        totalColIndex = h + 1;
        break;
      }
    }
    // If column doesn't exist, create it
    if (totalColIndex < 0) {
      totalColIndex = lastCol + 1;
      sheet.getRange(1, totalColIndex).setValue('应付总额 Total Due');
    }
    // Write the total to the last row
    sheet.getRange(lastRow, totalColIndex).setValue('$' + totalDue);
  } catch (err) {
    Logger.log('Error writing to sheet: ' + err);
  }

  // Send confirmation email
  if (email) {
    var subject = '谷雨暑期选修课报名确认 GR EDU Summer Electives Registration Confirmation';
    var body =
      '亲爱的 ' + parentName + '，\n' +
      'Dear ' + parentName + ',\n\n' +
      '感谢您为 ' + studentName + ' 报名谷雨暑期选修课！\n' +
      'Thank you for registering ' + studentName + ' for GR EDU Summer Electives!\n\n' +
      '━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
      '报名信息 Registration Summary\n' +
      '━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
      '学生 Student: ' + studentName + '\n' +
      '年级 Grade: ' + grade + '\n' +
      '校区 Campus: ' + campus + '\n' +
      'Sessions: ' + sessionCount + ' session(s)\n\n' +
      '选课 Courses:\n';

    if (hasChinese) body += '  • ' + chineseLevel + '\n';
    if (hasEnglish) body += '  • ' + englishLevel + '\n';
    if (hasMath) body += '  • ' + mathLevel + '\n';

    body += '\n━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
      '学费明细 Tuition Breakdown\n' +
      '━━━━━━━━━━━━━━━━━━━━━━━━━━\n';

    for (var b = 0; b < breakdown.length; b++) {
      body += '  ' + breakdown[b] + '\n';
    }
    body += '  × ' + sessionCount + ' session(s)\n';
    body += '\n  ★ 应付总额 Total Due: $' + totalDue + '\n';

    body += '\n━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
      '缴费方式 Payment\n' +
      '━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
      'Zelle 转账至 Send to: gredu2019@gmail.com\n' +
      '备注 Memo: ' + studentName + ' - Summer Elective\n\n' +
      '请尽快完成付款，名额以付款先后顺序确认。\n' +
      'Please complete payment promptly. Spots confirmed on a first-paid basis.\n\n' +
      '如有问题请回复此邮件。\n' +
      'Reply to this email if you have any questions.\n\n' +
      '谷雨教育 GR EDU';

    GmailApp.sendEmail(email, subject, body);
  }
}
