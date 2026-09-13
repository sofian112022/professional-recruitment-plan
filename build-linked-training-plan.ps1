$ErrorActionPreference = "Stop"
$out = Join-Path (Get-Location) "خطة-تدريب-جهار-احترافية.xlsx"
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Add()

function C($s,$r,$c,$v) {
  try { $s.Cells.Item($r,$c) = $v }
  catch { Write-Error ("Cell write failed: sheet={0}, row={1}, col={2}, value={3}" -f $s.Name,$r,$c,$v); throw }
}
function T($s,$n,$text) {
  $s.Range($s.Cells.Item(1,1),$s.Cells.Item(1,$n)).Merge()
  C $s 1 1 $text
  $x=$s.Range($s.Cells.Item(1,1),$s.Cells.Item(1,$n))
  $x.Font.Bold=$true; $x.Font.Size=16; $x.Font.Color=0xFFFFFF
  $x.Interior.Color=0x17365D; $x.HorizontalAlignment=-4108
}
function FH($s,$r,$n) {
  $x=$s.Range($s.Cells.Item($r,1),$s.Cells.Item($r,$n))
  $x.Font.Bold=$true; $x.Font.Color=0xFFFFFF; $x.Interior.Color=0x1F4E78
  $x.WrapText=$true; $x.HorizontalAlignment=-4108
}
function Table($s,$name,$lastRow,$lastCol) {
  $t=$s.ListObjects.Add(1,$s.Range($s.Cells.Item(3,1),$s.Cells.Item($lastRow,$lastCol)),$null,1)
  $t.Name=$name; $t.TableStyle="TableStyleMedium2"
}
function Finish($s) {
  $s.DisplayRightToLeft=$true; $s.Cells.Font.Name="Arial"; $s.Cells.Font.Size=10
  $u=$s.UsedRange; $u.WrapText=$true; $u.VerticalAlignment=-4108
  $u.Borders.LineStyle=1; $u.Borders.Weight=2
  $u.Columns.AutoFit() | Out-Null; $u.Rows.AutoFit() | Out-Null
}
function Headers($s,$headers) {
  for($i=0;$i -lt $headers.Count;$i++){C $s 3 ($i+1) $headers[$i]}
  FH $s 3 $headers.Count
}
function Rows($s,$data,[int]$start=4) {
  for($r=0;$r -lt $data.Count;$r++){for($c=0;$c -lt $data[$r].Count;$c++){C $s ($start+$r) ($c+1) $data[$r][$c]}}
}

try {
  $setup=$wb.Worksheets.Item(1); $setup.Name="Setup"
  T $setup 4 "إعدادات النظام وتعريف خطة التدريب"
  Headers $setup @("البند","القيمة","شرح الاستخدام","مصدر/مسؤول الاعتماد")
  Rows $setup @(
    @("اسم المنشأة","يُستكمل","الاسم الرسمي للمنشأة","إدارة المنشأة"),
    @("نوع المنشأة","مستشفى","اختر النوع الفعلي","الجودة"),
    @("سنة الخطة",2026,"غيّر السنة عند بدء دورة جديدة","الموارد البشرية"),
    @("نسبة النجاح",70,"الحد الأدنى للتقييم البعدي","الجودة والتدريب"),
    @("نسبة الحضور",80,"الحد الأدنى لإكمال النشاط","الجودة والتدريب"),
    @("إصدار الدليل","يُستكمل","سجل رقم وتاريخ آخر دليل GAHAR معتمد","الجودة"),
    @("تنبيه","هذه أداة تشغيلية وليست بديلًا عن دليل GAHAR الرسمي","طابق البنود مع نطاق المنشأة والإصدار الساري قبل الاعتماد","مسؤول الجودة")
  )
  $setup.Columns.Item(1).ColumnWidth=24;$setup.Columns.Item(2).ColumnWidth=42;$setup.Columns.Item(3).ColumnWidth=60;$setup.Columns.Item(4).ColumnWidth=24
  Finish $setup

  $stand=$wb.Worksheets.Add();$stand.Name="Standards";T $stand 7 "كتالوج محاور ومعايير التدريب - مرجع التخصيص"
  Headers $stand @("كود المحور","المحور","موضوع/متطلب التدريب","الفئات المعنية","دليل الإثبات المتوقع","تكرار مقترح","ملاحظات المطابقة")
  Rows $stand @(
    @("HRP","إدارة الموارد البشرية","التعريف، الوصف الوظيفي، الكفاءة، التقييم وإعادة التدريب","كل العاملين","خطة تدريب، مصفوفة كفاءة، سجلات حضور وتقييم","سنوي/عند التعيين","اربطه بسياسة المنشأة"),
    @("PFR","حقوق المريض والأسرة","الحقوق، الخصوصية، الموافقة، الشكاوى والتواصل","كل العاملين","مادة، اختبار، حالات تطبيقية، سجل شكاوى","سنوي","تحقق من سياسة حقوق المريض"),
    @("IPC","مكافحة العدوى","نظافة اليدين، PPE، العزل، النفايات، التعرض المهني","سريري/خدمات معاونة","قائمة ملاحظة، تدريب عملي، نتائج تدقيق","ربع سنوي","حسب تقييم مخاطر العدوى"),
    @("MMU","إدارة الدواء","التخزين، التحقق، الأدوية عالية الخطورة، الإبلاغ","أطباء/تمريض/صيدلة","محاكاة، اختبار، مراجعة خطأ دوائي","سنوي","حسب قائمة الأدوية الحرجة"),
    @("FMS","السلامة وإدارة المنشأة","الحريق، الإخلاء، الطوارئ، المواد الخطرة، الأمن","كل العاملين","سيناريو، زمن إخلاء، محضر تمرين","سنوي/نصف سنوي","حسب خطة الطوارئ"),
    @("QPS","الجودة وسلامة المريض","الإبلاغ، تحليل السبب، مؤشرات الأداء، PDSA","قادة الأقسام والجودة","خطة تحسين، محضر لجنة، مؤشر قبل/بعد","ربع سنوي","اربطه بالمخاطر الفعلية"),
    @("HIS","المعلومات والسجلات","التوثيق، السرية، الصلاحيات، الأمن السيبراني","سريري/سجلات/IT","تدقيق ملفات، اختبار، سجل صلاحيات","سنوي","طبق سياسة حماية البيانات"),
    @("GLD","القيادة والحوكمة","المسؤوليات، إدارة المخاطر، استمرارية الأعمال","القيادة ورؤساء الأقسام","محاضر، تمرين، خطة استمرارية","سنوي","وفق هيكل المنشأة")
  );Table $stand "tblStandards" 11 7;Finish $stand

  $emp=$wb.Worksheets.Add();$emp.Name="Employees";T $emp 12 "قاعدة بيانات العاملين - الإدخال الأساسي"
  Headers $emp @("كود الموظف","الاسم","القسم","المسمى","الفئة","تاريخ التعيين","الحالة","المدير","عدد الأنشطة","نسبة الحضور","متوسط التقييم","حالة التدريب")
  Rows $emp @(
    @("EMP-001","يُستكمل","التمريض","ممرض/ة","سريري","01/01/2025","نشط","رئيس التمريض","","","",""),
    @("EMP-002","يُستكمل","الصيدلية","صيدلي","سريري","15/02/2025","نشط","مدير الصيدلية","","","",""),
    @("EMP-003","يُستكمل","الاستقبال","موظف استقبال","إداري","10/03/2025","نشط","مدير الخدمة","","","","")
  )
  for($r=4;$r -le 203;$r++){
    C $emp $r 9 "=IF(A$r="""","""",COUNTIF(Attendance!D:D,A$r))"
    C $emp $r 10 "=IFERROR(COUNTIFS(Attendance!D:D,A$r,Attendance!F:F,""نعم"")/I$r,0)"
    C $emp $r 11 "=IFERROR(AVERAGEIF(Evaluation!D:D,A$r,Evaluation!H:H),"""")"
    C $emp $r 12 "=IF(A$r="""","""",IF(J$r<Setup!B8,""حضور غير مكتمل"",IF(K$r<Setup!B7,""يحتاج إعادة تدريب"",""مكتمل"")))"
  }
  $emp.Range("G4:G203").Validation.Add(3,1,1,"نشط,موقوف,منتهية الخدمة");$emp.Range("G4:G203").Validation.InCellDropdown=$true
  $emp.Range("J4:K203").NumberFormat="0%";Table $emp "tblEmployees" 203 12;Finish $emp

  $needs=$wb.Worksheets.Add();$needs.Name="Needs";T $needs 11 "تحليل الاحتياج التدريبي - قبل إعداد الخطة"
  Headers $needs @("كود الاحتياج","القسم","الفجوة/الخطر","مصدر الاحتياج","الأثر","الأولوية","الكفاءة المطلوبة","الإجراء التدريبي","المسؤول","موعد الإغلاق","الحالة")
  Rows $needs @(
    @("NEED-001","كل الأقسام","عدم اكتمال سجلات التدريب والكفاءة","تدقيق داخلي","مرتفع","عاجل","إدارة التدريب والتوثيق","بناء النظام وتدريب المستخدمين","الموارد البشرية","31/01/2026","مفتوح"),
    @("NEED-002","التمريض","تفاوت الالتزام بنظافة اليدين","مؤشر مكافحة العدوى","مرتفع","عاجل","مكافحة العدوى","تدريب عملي وتدقيق ملاحظة","مكافحة العدوى","28/02/2026","مفتوح"),
    @("NEED-003","الأدوية","حاجة إلى توحيد التحقق من الدواء","بلاغات/مخاطر","مرتفع","عاجل","سلامة الدواء","محاكاة واختبار كفاءة","الصيدلية","31/03/2026","مفتوح")
  );$needs.Range("F4:F203").Validation.Add(3,1,1,"عاجل,مرتفع,متوسط,منخفض");$needs.Range("K4:K203").Validation.Add(3,1,1,"مفتوح,قيد التنفيذ,مغلق");Table $needs "tblNeeds" 203 11;Finish $needs

  $plan=$wb.Worksheets.Add();$plan.Name="Plan";T $plan 17 "الخطة السنوية المعتمدة - اربط كل نشاط باحتياج ومحور ودليل"
  Headers $plan @("كود النشاط","كود الاحتياج","كود المحور","الشهر","عنوان النشاط","الهدف القابل للقياس","الفئة","نوع النشاط","المسؤول","الساعات","التاريخ المخطط","التاريخ الفعلي","الحالة","عدد الحضور","متوسط التقييم","دليل الإثبات","ملاحظات التحسين")
  $pdata=@(
    @("TR-001","NEED-001","HRP","يناير","نظام التدريب ومصفوفة الكفاءة","يسجل المستخدم نشاطًا وموظفًا وتقييمًا دون أخطاء","مسؤولو التدريب","ورشة نظام","الموارد البشرية",3,"15/01/2026","","مخطط","","","",""),
    @("TR-002","NEED-001","HRP","يناير","حقوق وواجبات العاملين ومدونة السلوك","يشرح العامل 5 حقوق و5 واجبات وقناة إبلاغ","كل العاملين","تعريفي","الموارد البشرية",2,"22/01/2026","","مخطط","","","",""),
    @("TR-003","NEED-002","IPC","فبراير","نظافة اليدين ومعدات الوقاية","يجتاز العامل قائمة ملاحظة المهارة بنسبة 80%","التمريض والخدمات","عملي","مكافحة العدوى",4,"12/02/2026","","مخطط","","","",""),
    @("TR-004","NEED-003","MMU","مارس","سلامة الدواء والأدوية عالية الخطورة","يطبق خطوات التحقق الخمس في محاكاة","أطباء وتمريض وصيدلة","محاكاة","مدير الصيدلية",3,"10/03/2026","","مخطط","","","",""),
    @("TR-005","","FMS","مارس","الحريق والإخلاء ونقطة التجمع","يصل الفريق لنقطة التجمع خلال الزمن المستهدف","كل العاملين","تمرين طوارئ","السلامة المهنية",3,"24/03/2026","","مخطط","","","",""),
    @("TR-006","","QPS","أبريل","الإبلاغ عن الحوادث وتحليل السبب الجذري","يبني الفريق خطة PDSA لحادثة تدريبية","قادة الأقسام","دراسة حالة","الجودة",3,"14/04/2026","","مخطط","","","",""),
    @("TR-007","","HIS","مايو","التوثيق الطبي وسرية المعلومات","يحقق الملف التدريبي 90% في قائمة التدقيق","سريري وسجلات","تطبيقي","السجلات الطبية",3,"12/05/2026","","مخطط","","","",""),
    @("TR-008","","PFR","يونيو","حقوق المريض والشكاوى والموافقة","يتعامل مع الحالة وفق السياسة دون خرق خصوصية","كل العاملين","لعب أدوار","خدمة العملاء",2,"16/06/2026","","مخطط","","","",""),
    @("TR-009","","AOP/COP","يوليو","التقييم الأولي والتصعيد والتسليم SBAR","ينفذ تسليم حالة منظمًا بدرجة 80%","أطباء وتمريض","محاكاة","الإدارة الطبية",4,"14/07/2026","","مخطط","","","",""),
    @("TR-010","","FMS","أغسطس","النفايات والمواد الخطرة والتعرض المهني","يفرز وينقل النفايات وفق القائمة المعتمدة","كل الفئات المعنية","عملي","السلامة والعدوى",2,"11/08/2026","","مخطط","","","",""),
    @("TR-011","","HRP","سبتمبر","تقييم الكفاءة وإعادة التدريب","يحدد المدير فجوة الكفاءة ويضع إجراءً فرديًا","رؤساء الأقسام","تقييم عملي","الموارد البشرية",4,"15/09/2026","","مخطط","","","",""),
    @("TR-012","","GLD/QPS","أكتوبر","إدارة الأزمات واستمرارية الأعمال","يحدد الفريق الأدوار ومسارات التصعيد","القيادة والطوارئ","تمرين مكتبي","إدارة المنشأة",3,"13/10/2026","","مخطط","","","",""),
    @("TR-013","","QPS/HRP","ديسمبر","المراجعة السنوية وخطة العام القادم","يعرض المسؤول مؤشرات التدريب وخطة تحسين","الإدارة والجودة","مراجعة","الجودة والموارد البشرية",3,"08/12/2026","","مخطط","","","","")
  )
  Rows $plan $pdata
  for($r=4;$r -le 203;$r++){C $plan $r 14 "=IF(A$r="""","""",COUNTIFS(Attendance!B:B,A$r,Attendance!F:F,""نعم""))";C $plan $r 15 "=IFERROR(AVERAGEIF(Evaluation!B:B,A$r,Evaluation!H:H),"""")"}
  $plan.Range("M4:M203").Validation.Add(3,1,1,"مخطط,منفذ,مؤجل,ملغى");$plan.Range("M4:M203").Validation.InCellDropdown=$true
  $plan.Range("P4:P203").NumberFormat="@";Table $plan "tblPlan" 203 17;Finish $plan

  $sessions=$wb.Worksheets.Add();$sessions.Name="Sessions";T $sessions 10 "سجل التنفيذ الفعلي - جلسة لكل تنفيذ"
  Headers $sessions @("كود الجلسة","كود النشاط","التاريخ الفعلي","المكان","المدرب","السعة","عدد الحضور","نسبة الحضور","التكلفة","حالة الجلسة")
  Rows $sessions @(@("SES-001","TR-001","15/01/2026","قاعة التدريب","الموارد البشرية",25,"","","","مخطط"),@("SES-002","TR-002","22/01/2026","قاعة التدريب","الموارد البشرية",30,"","","","مخطط"))
  for($r=4;$r -le 203;$r++){C $sessions $r 7 "=IF(A$r="""","""",COUNTIFS(Attendance!B:B,B$r,Attendance!J:J,C$r,Attendance!F:F,""نعم""))";C $sessions $r 8 "=IFERROR(G$r/F$r,0)"}
  $sessions.Range("J4:J203").Validation.Add(3,1,1,"مخطط,منفذ,ملغى");$sessions.Range("H4:H203").NumberFormat="0%"
  Table $sessions "tblSessions" 203 10;Finish $sessions

  $att=$wb.Worksheets.Add();$att.Name="Attendance";T $att 12 "الحضور - الإدخال اليومي الوحيد للحضور"
  Headers $att @("رقم","كود الجلسة","كود النشاط","التاريخ","كود الموظف","اسم الموظف","القسم","حضر؟","وقت الدخول","توقيع/إثبات","حالة السجل","ملاحظات")
  for($r=4;$r -le 1003;$r++){C $att $r 1 ($r-3);C $att $r 3 "=IFERROR(VLOOKUP(B$r,Sessions!A:J,2,FALSE),"""")";C $att $r 4 "=IFERROR(VLOOKUP(B$r,Sessions!A:J,3,FALSE),"""")";C $att $r 6 "=IFERROR(VLOOKUP(E$r,Employees!A:L,2,FALSE),"""")";C $att $r 7 "=IFERROR(VLOOKUP(E$r,Employees!A:L,3,FALSE),"""")";C $att $r 11 "=IF(B$r="""","""",IF(OR(E$r="""",H$r=""""),""ناقص"",IF(H$r=""نعم"",""مكتمل"",""غائب"")))"}
  $att.Range("B4:B1003").Validation.Add(3,1,1,"SES-001,SES-002");$att.Range("B4:B1003").Validation.InCellDropdown=$true
  $att.Range("E4:E1003").Validation.Add(3,1,1,"EMP-001,EMP-002,EMP-003");$att.Range("E4:E1003").Validation.InCellDropdown=$true
  $att.Range("H4:H1003").Validation.Add(3,1,1,"نعم,لا");$att.Range("H4:H1003").Validation.InCellDropdown=$true
  Table $att "tblAttendance" 1003 12;Finish $att

  $eval=$wb.Worksheets.Add();$eval.Name="Evaluation";T $eval 13 "التقييم - أدخل الدرجات وتظهر النتيجة تلقائيًا"
  Headers $eval @("رقم","كود الجلسة","كود النشاط","كود الموظف","اسم الموظف","قبلي %","بعدي %","الدرجة المعتمدة","النتيجة","مقيم الكفاءة","تاريخ التقييم","إجراء التحسين","رابط نموذج التقييم")
  for($r=4;$r -le 1003;$r++){C $eval $r 1 ($r-3);C $eval $r 3 "=IFERROR(VLOOKUP(B$r,Sessions!A:J,2,FALSE),"""")";C $eval $r 5 "=IFERROR(VLOOKUP(D$r,Employees!A:L,2,FALSE),"""")";C $eval $r 8 "=IF(G$r="""","""",G$r)";C $eval $r 9 "=IF(H$r="""","""",IF(H$r>=Setup!B7,""ناجح"",""إعادة تدريب""))"}
  $eval.Range("B4:B1003").Validation.Add(3,1,1,"SES-001,SES-002");$eval.Range("D4:D1003").Validation.Add(3,1,1,"EMP-001,EMP-002,EMP-003")
  $eval.Range("F4:H1003").NumberFormat="0%";Table $eval "tblEvaluation" 1003 13;Finish $eval

  $comp=$wb.Worksheets.Add();$comp.Name="Competency";T $comp 12 "مصفوفة الكفاءة - القرار النهائي لكل موظف"
  Headers $comp @("كود الموظف","اسم الموظف","القسم","الكفاءة/المهارة","النشاط المرجعي","درجة الكفاءة %","تاريخ التقييم","المقيم","الحالة","موعد إعادة التقييم","إجراء تصحيحي","رابط الدليل")
  Rows $comp @(
    @("EMP-001","","","نظافة اليدين","TR-003","","","","لم يقيم","","",""),
    @("EMP-002","","","سلامة الدواء","TR-004","","","","لم يقيم","","",""),
    @("EMP-003","","","حقوق المريض والشكاوى","TR-008","","","","لم يقيم","","","")
  )
  for($r=4;$r -le 203;$r++){C $comp $r 2 "=IFERROR(VLOOKUP(A$r,Employees!A:L,2,FALSE),"""")";C $comp $r 3 "=IFERROR(VLOOKUP(A$r,Employees!A:L,3,FALSE),"""")";C $comp $r 9 "=IF(F$r="""",""لم يقيم"",IF(F$r>=Setup!B7,""كفء"",""غير كفء - إعادة تدريب""))"}
  $comp.Range("F4:F203").NumberFormat="0%";$comp.Range("I4:I203").Validation.Add(3,1,1,"لم يقيم,كفء,غير كفء - إعادة تدريب");Table $comp "tblCompetency" 203 12;Finish $comp

  $ev=$wb.Worksheets.Add();$ev.Name="Evidence";T $ev 10 "سجل الأدلة - اربط كل نشاط بملف يمكن عرضه أثناء المراجعة"
  Headers $ev @("كود الدليل","كود النشاط","كود الجلسة","نوع الدليل","اسم الملف/الرابط","المسؤول","تاريخ الحفظ","الحالة","مراجع الجودة","ملاحظات")
  Rows $ev @(
    @("EVD-001","TR-001","SES-001","مادة تدريب","","مسؤول التدريب","15/01/2026","مطلوب","",""),
    @("EVD-002","TR-001","SES-001","كشف حضور","","مسؤول التدريب","15/01/2026","مطلوب","",""),
    @("EVD-003","TR-001","SES-001","نتائج تقييم وخطة تحسين","","مسؤول التدريب","15/01/2026","مطلوب","","")
  )
  $ev.Range("H4:H203").Validation.Add(3,1,1,"مطلوب,مكتمل,قيد المراجعة,غير منطبق");Table $ev "tblEvidence" 203 10;Finish $ev

  $dash=$wb.Worksheets.Add();$dash.Name="Dashboard";T $dash 5 "لوحة مؤشرات الإدارة والجودة"
  Headers $dash @("المؤشر","القيمة","المستهدف","الحالة","مصدر البيانات")
  $m=@(
    @("الموظفون النشطون","=COUNTIF(Employees!G:G,""نشط"")", "كل الموظفين","","Employees"),
    @("الأنشطة المخططة","=COUNTIF(Plan!A:A,""TR-*"")","حسب الخطة","","Plan"),
    @("الأنشطة المنفذة","=COUNTIF(Plan!M:M,""منفذ"")","100% من المستحق","","Plan"),
    @("نسبة تنفيذ الخطة","=IFERROR(B7/B6,0)", ">=90%","","Plan"),
    @("سجلات الحضور","=COUNTIF(Attendance!E:E,""EMP-*"")","حسب الجلسات","","Attendance"),
    @("نسبة الحضور","=IFERROR(COUNTIF(Attendance!H:H,""نعم"")/B8,0)",">=80%","","Attendance"),
    @("متوسط التقييم البعدي","=IFERROR(AVERAGE(Evaluation!H:H),0)",">=70%","","Evaluation"),
    @("حالات إعادة التدريب","=COUNTIF(Evaluation!I:I,""إعادة تدريب"")","0 أو خطة تصحيح","","Evaluation"),
    @("أدلة مكتملة","=COUNTIF(Evidence!H:H,""مكتمل"")","100% للأنشطة المنفذة","","Evidence"),
    @("كفاءات غير مكتملة","=COUNTIF(Competency!I:I,""غير كفء - إعادة تدريب"")","0 أو خطة تصحيح","","Competency")
  )
  Rows $dash $m
  for($r=4;$r -le 13;$r++){C $dash $r 4 "=IF(B$r="""","""",IF(ISNUMBER(B$r),IF(OR($C$r=""0 أو خطة تصحيح"",B$r>=0),""راجع المؤشر"",""راجع المؤشر""),""""))"}
  $dash.Range("B7").NumberFormat="0%";$dash.Range("B9:B10").NumberFormat="0%"
  $dash.Columns.Item(1).ColumnWidth=30;$dash.Columns.Item(2).ColumnWidth=18;$dash.Columns.Item(3).ColumnWidth=20;$dash.Columns.Item(4).ColumnWidth=18;$dash.Columns.Item(5).ColumnWidth=18;Finish $dash

  $monthly=$wb.Worksheets.Add();$monthly.Name="Monthly";T $monthly 8 "التقرير الشهري - يختار المستخدم الشهر"
  C $monthly 3 1 "الشهر";C $monthly 3 2 "يناير";C $monthly 4 1 "عدد الأنشطة";C $monthly 4 2 '=COUNTIF(Plan!D:D,B3)'
  C $monthly 5 1 "المنفذ";C $monthly 5 2 '=COUNTIFS(Plan!D:D,B3,Plan!M:M,"منفذ")'
  C $monthly 6 1 "نسبة التنفيذ";C $monthly 6 2 '=IFERROR(B5/B4,0)'
  C $monthly 7 1 "عدد الحضور";C $monthly 7 2 "=SUMPRODUCT(--(TEXT(Attendance!D4:D1003,""mmmm"")=B3),--(Attendance!H4:H1003=""نعم""))"
  C $monthly 8 1 "الإجراءات المفتوحة";C $monthly 8 2 '=COUNTIF(Needs!K:K,"مفتوح")'
  $monthly.Range("B3").Validation.Add(3,1,1,"يناير,فبراير,مارس,أبريل,مايو,يونيو,يوليو,أغسطس,سبتمبر,أكتوبر,نوفمبر,ديسمبر");$monthly.Range("B3").Validation.InCellDropdown=$true;$monthly.Range("B6").NumberFormat="0%";$monthly.Columns.Item(1).ColumnWidth=28;$monthly.Columns.Item(2).ColumnWidth=20;Finish $monthly

  foreach($s in $wb.Worksheets){$s.Activate();$s.Application.ActiveWindow.SplitRow=3;$s.Application.ActiveWindow.FreezePanes=$true}
  $setup.Activate();$wb.SaveAs($out,51);Write-Output $out
}
finally {
  if($wb){$wb.Close($true)};if($xl){$xl.Quit()}
  [System.GC]::Collect();[System.GC]::WaitForPendingFinalizers()
}
