$ErrorActionPreference = "Stop"

$output = Join-Path (Get-Location) "خطة-تدريب-سنوية-مترابطة.xlsx"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Add()

function Cell($s, [int]$r, [int]$c, $v) { $s.Cells.Item($r, $c) = $v }
function Title($s, [int]$last, [string]$text) {
    $s.Range($s.Cells.Item(1,1), $s.Cells.Item(1,$last)).Merge()
    Cell $s 1 1 $text
    $x = $s.Range($s.Cells.Item(1,1), $s.Cells.Item(1,$last))
    $x.Font.Bold = $true; $x.Font.Size = 16; $x.Font.Color = 0xFFFFFF
    $x.Interior.Color = 0x1F4E78; $x.HorizontalAlignment = -4108
}
function Header($s, [int]$r, [int]$last) {
    $x = $s.Range($s.Cells.Item($r,1), $s.Cells.Item($r,$last))
    $x.Font.Bold = $true; $x.Font.Color = 0xFFFFFF; $x.Interior.Color = 0x5B9BD5
    $x.WrapText = $true; $x.HorizontalAlignment = -4108
}
function Finish($s) {
    $s.DisplayRightToLeft = $true
    $s.Cells.Font.Name = "Arial"; $s.Cells.Font.Size = 10
    $u = $s.UsedRange; $u.WrapText = $true; $u.VerticalAlignment = -4108
    $u.Borders.LineStyle = 1; $u.Borders.Weight = 2
    $u.Columns.AutoFit() | Out-Null; $u.Rows.AutoFit() | Out-Null
}
function Add-Table($s, [string]$name, [int]$lastRow, [int]$lastCol) {
    $range = $s.Range($s.Cells.Item(3,1), $s.Cells.Item($lastRow,$lastCol))
    $table = $s.ListObjects.Add(1, $range, $null, 1)
    $table.Name = $name
    $table.TableStyle = "TableStyleMedium2"
}

try {
    $settings = $wb.Worksheets.Item(1); $settings.Name = "الإعدادات"
    Title $settings 4 "نظام خطة التدريب والتقييم المترابط"
    $settingsRows = @(
        @("اسم المنشأة", "يُستكمل"),
        @("نوع المنشأة", "مستشفى / مركز طبي"),
        @("سنة الخطة", "2026"),
        @("حد النجاح %", 70),
        @("حد الحضور %", 80),
        @("مسؤول التدريب", "يُستكمل"),
        @("ملاحظة", "راجع آخر إصدار رسمي من دليل GAHAR ونطاق المنشأة قبل الاعتماد.")
    )
    Cell $settings 3 1 "البند"; Cell $settings 3 2 "القيمة"; Header $settings 3 2
    for ($i=0; $i -lt $settingsRows.Count; $i++) { Cell $settings ($i+4) 1 $settingsRows[$i][0]; Cell $settings ($i+4) 2 $settingsRows[$i][1] }
    $settings.Columns.Item(1).ColumnWidth = 24; $settings.Columns.Item(2).ColumnWidth = 70
    Finish $settings

    $employees = $wb.Worksheets.Add(); $employees.Name = "الموظفون"
    Title $employees 12 "قاعدة بيانات الموظفين"
    $eh = @("كود الموظف","اسم الموظف","القسم","المسمى الوظيفي","نوع التعاقد","تاريخ التعيين","حالة الموظف","المدير المباشر","الهاتف/البريد","عدد الأنشطة","نسبة الإكمال","ملاحظات")
    for ($c=0; $c -lt $eh.Count; $c++) { Cell $employees 3 ($c+1) $eh[$c] }; Header $employees 3 $eh.Count
    $er = @(
        @("EMP-001","أحمد محمد","التمريض","ممرض/ة","دوام كامل","01/01/2025","نشط","رئيس التمريض","","=COUNTIF(الحضور!D:D,A4)","=IFERROR(COUNTIFS(الحضور!D:D,A4,الحضور!F:F,""نعم"")/J4,0)",""),
        @("EMP-002","سارة علي","الصيدلية","صيدلي","دوام كامل","15/02/2025","نشط","مدير الصيدلية","","=COUNTIF(الحضور!D:D,A5)","=IFERROR(COUNTIFS(الحضور!D:D,A5,الحضور!F:F,""نعم"")/J5,0)",""),
        @("EMP-003","محمود حسن","الاستقبال","موظف استقبال","دوام كامل","10/03/2025","نشط","مدير خدمة العملاء","","=COUNTIF(الحضور!D:D,A6)","=IFERROR(COUNTIFS(الحضور!D:D,A6,الحضور!F:F,""نعم"")/J6,0)","")
    )
    for ($r=0; $r -lt $er.Count; $r++) { for ($c=0; $c -lt $er[$r].Count; $c++) { Cell $employees ($r+4) ($c+1) $er[$r][$c] } }
    for ($r=7; $r -le 103; $r++) { Cell $employees $r 10 ("=IF(A$r="""","""",COUNTIF(الحضور!D:D,A$r))"); Cell $employees $r 11 ("=IFERROR(COUNTIFS(الحضور!D:D,A$r,الحضور!F:F,""نعم"")/J$r,0)") }
    $employees.Range("G4:G103").Validation.Add(3,1,1,"نشط,موقوف,منتهية الخدمة"); $employees.Range("G4:G103").Validation.InCellDropdown = $true
    $employees.Range("K4:K103").NumberFormat = "0%"
    Add-Table $employees "tblEmployees" 103 12; Finish $employees

    $plan = $wb.Worksheets.Add(); $plan.Name = "الخطة"
    Title $plan 15 "الخطة السنوية للتدريب"
    $ph = @("كود النشاط","الشهر","محور GAHAR","موضوع التدريب","الفئة المستهدفة","الهدف والكفاءة","نوع التدريب","المسؤول","الساعات","التاريخ المخطط","التاريخ المنفذ","الحالة","عدد الحضور","متوسط التقييم","رابط الدليل")
    for ($c=0; $c -lt $ph.Count; $c++) { Cell $plan 3 ($c+1) $ph[$c] }; Header $plan 3 $ph.Count
    $pr = @(
        @("TR-001","يناير","HRP / GLD","التعريف بالمنشأة وحقوق وواجبات العاملين","جميع العاملين","فهم السياسات ومدونة السلوك وقنوات الإبلاغ","تعريفي","الموارد البشرية",3,"15/01","","مخطط","","",""),
        @("TR-002","فبراير","IPC","الوقاية من العدوى ومكافحة العدوى","السريريون والخدمات المعاونة","تطبيق نظافة اليدين ومعدات الوقاية والعزل","عملي","فريق مكافحة العدوى",4,"12/02","","مخطط","","",""),
        @("TR-003","مارس","MMU","سلامة الدواء والإبلاغ عن الأخطاء","الأطباء والتمريض والصيدلة","تطبيق التحقق الآمن من الدواء والإبلاغ","محاكاة","مدير الصيدلية",3,"10/03","","مخطط","","",""),
        @("TR-004","مارس","FMS","السلامة من الحريق والإخلاء","جميع العاملين","تنفيذ الإنذار والإخلاء واستخدام الطفاية","تمرين","السلامة المهنية",3,"24/03","","مخطط","","",""),
        @("TR-005","مايو","QPS","سلامة المريض والتحسين المستمر","قادة الوحدات والجودة","تحليل السبب الجذري وبناء خطة PDSA","دراسة حالة","إدارة الجودة",3,"12/05","","مخطط","","",""),
        @("TR-006","يونيو","HIS","التوثيق الطبي وسرية المعلومات","الأطباء والتمريض والسجلات","تسجيل دقيق وحماية البيانات والصلاحيات","تطبيقي","السجلات الطبية",3,"16/06","","مخطط","","",""),
        @("TR-007","يوليو","AOP / COP","التقييم الأولي والتصعيد السريري","الأطباء والتمريض والاستقبال","اكتشاف التدهور والتواصل المنظم SBAR","محاكاة","الإدارة الطبية",4,"14/07","","مخطط","","",""),
        @("TR-008","سبتمبر","PFR / QPS","تجربة المريض والشكاوى","الاستقبال وخدمة العملاء","استقبال الشكوى والرد والتحسين","لعب أدوار","خدمة العملاء",2,"15/09","","مخطط","","",""),
        @("TR-009","أكتوبر","HRP","تقييم الكفاءة وإعادة التأهيل","جميع الفئات","إثبات الكفاءة وتحديد فجوات الأداء","تقييم عملي","الموارد البشرية",4,"13/10","","مخطط","","",""),
        @("TR-010","نوفمبر","GLD / QPS","إدارة الأزمات واستمرارية الأعمال","لجنة الطوارئ وقادة الوحدات","تنفيذ الاستجابة واستمرارية الخدمات","تمرين مكتبي","إدارة المنشأة",3,"17/11","","مخطط","","",""),
        @("TR-011","ديسمبر","QPS / HRP","المراجعة السنوية وخطة العام القادم","الإدارة ورؤساء الأقسام","تحليل مؤشرات التدريب وخطة التحسين","مراجعة","الجودة والموارد البشرية",3,"08/12","","مخطط","","","")
    )
    for ($r=0; $r -lt $pr.Count; $r++) { for ($c=0; $c -lt $pr[$r].Count; $c++) { Cell $plan ($r+4) ($c+1) $pr[$r][$c] } }
    for ($r=4; $r -le 103; $r++) {
        Cell $plan $r 13 ("=IF(A$r="""","""",COUNTIFS(الحضور!B:B,A$r,الحضور!F:F,""نعم""))")
        Cell $plan $r 14 ("=IFERROR(AVERAGEIF(التقييمات!B:B,A$r,التقييمات!H:H),"""")")
    }
    $plan.Range("L4:L103").Validation.Add(3,1,1,"مخطط,منفذ,مؤجل,ملغى"); $plan.Range("L4:L103").Validation.InCellDropdown = $true
    Add-Table $plan "tblPlan" 103 15; Finish $plan

    $attendance = $wb.Worksheets.Add(); $attendance.Name = "الحضور"
    Title $attendance 12 "سجل الحضور المركزي - أدخل البيانات هنا"
    $ah = @("رقم السجل","كود النشاط","موضوع التدريب","كود الموظف","اسم الموظف","حضر؟","وقت الحضور","التوقيع/الإثبات","المدرب","التاريخ","حالة السجل","ملاحظات")
    for ($c=0; $c -lt $ah.Count; $c++) { Cell $attendance 3 ($c+1) $ah[$c] }; Header $attendance 3 $ah.Count
    for ($r=4; $r -le 503; $r++) {
        Cell $attendance $r 1 $r-3
        Cell $attendance $r 3 ("=IFERROR(VLOOKUP(B$r,الخطة!A:O,4,FALSE),"""")")
        Cell $attendance $r 5 ("=IFERROR(VLOOKUP(D$r,الموظفون!A:L,2,FALSE),"""")")
        Cell $attendance $r 11 ("=IF(B$r="""","""",IF(OR(D$r="""",F$r=""""),""ناقص بيانات"",IF(F$r=""نعم"",""مكتمل"",""غائب"")))")
    }
    $attendance.Range("F4:F503").Validation.Add(3,1,1,"نعم,لا"); $attendance.Range("F4:F503").Validation.InCellDropdown = $true
    $attendance.Range("B4:B503").Validation.Add(3,1,1,"TR-001,TR-002,TR-003,TR-004,TR-005,TR-006,TR-007,TR-008,TR-009,TR-010,TR-011"); $attendance.Range("B4:B503").Validation.InCellDropdown = $true
    $attendance.Range("D4:D503").Validation.Add(3,1,1,"EMP-001,EMP-002,EMP-003"); $attendance.Range("D4:D503").Validation.InCellDropdown = $true
    Add-Table $attendance "tblAttendance" 503 12; Finish $attendance

    $scores = $wb.Worksheets.Add(); $scores.Name = "التقييمات"
    Title $scores 12 "سجل التقييمات - أدخل الدرجة والنتيجة"
    $sh = @("رقم التقييم","كود النشاط","موضوع التدريب","كود الموظف","اسم الموظف","تقييم قبلي %","تقييم بعدي %","الدرجة المعتمدة %","النتيجة","المقيّم","تاريخ التقييم","خطة التحسين")
    for ($c=0; $c -lt $sh.Count; $c++) { Cell $scores 3 ($c+1) $sh[$c] }; Header $scores 3 $sh.Count
    for ($r=4; $r -le 503; $r++) {
        Cell $scores $r 1 $r-3
        Cell $scores $r 3 ("=IFERROR(VLOOKUP(B$r,الخطة!A:O,4,FALSE),"""")")
        Cell $scores $r 5 ("=IFERROR(VLOOKUP(D$r,الموظفون!A:L,2,FALSE),"""")")
        Cell $scores $r 8 ("=IF(G$r="""","""",G$r)")
        Cell $scores $r 9 ("=IF(H$r="""","""",IF(H$r>=الإعدادات!B7,""ناجح"",""يحتاج إعادة تدريب""))")
    }
    $scores.Range("B4:B503").Validation.Add(3,1,1,"TR-001,TR-002,TR-003,TR-004,TR-005,TR-006,TR-007,TR-008,TR-009,TR-010,TR-011"); $scores.Range("B4:B503").Validation.InCellDropdown = $true
    $scores.Range("D4:D503").Validation.Add(3,1,1,"EMP-001,EMP-002,EMP-003"); $scores.Range("D4:D503").Validation.InCellDropdown = $true
    $scores.Range("F4:H503").NumberFormat = "0"
    Add-Table $scores "tblScores" 503 12; Finish $scores

    $dashboard = $wb.Worksheets.Add(); $dashboard.Name = "لوحة التحكم"
    Title $dashboard 4 "لوحة مؤشرات التدريب"
    $metrics = @(
        @("إجمالي الموظفين","=COUNTIF(الموظفون!G:G,""نشط"")"),
        @("إجمالي الأنشطة","=COUNTIF(الخطة!A:A,""TR-*"")"),
        @("الأنشطة المنفذة","=COUNTIF(الخطة!L:L,""منفذ"")"),
        @("نسبة تنفيذ الخطة","=IFERROR(B6/B5,0)"),
        @("إجمالي سجلات الحضور","=COUNTIF(الحضور!D:D,""EMP-*"")"),
        @("نسبة الحضور","=IFERROR(COUNTIF(الحضور!F:F,""نعم"")/B8,0)"),
        @("متوسط التقييم البعدي","=IFERROR(AVERAGE(التقييمات!H:H),0)"),
        @("عدد الناجحين","=COUNTIF(التقييمات!I:I,""ناجح"")"),
        @("يحتاجون إعادة تدريب","=COUNTIF(التقييمات!I:I,""يحتاج إعادة تدريب"")")
    )
    Cell $dashboard 3 1 "المؤشر"; Cell $dashboard 3 2 "القيمة"; Header $dashboard 3 2
    for ($i=0; $i -lt $metrics.Count; $i++) { Cell $dashboard ($i+4) 1 $metrics[$i][0]; Cell $dashboard ($i+4) 2 $metrics[$i][1] }
    $dashboard.Range("B7:B9").NumberFormat = "0%"
    $dashboard.Columns.Item(1).ColumnWidth = 32; $dashboard.Columns.Item(2).ColumnWidth = 22
    Finish $dashboard

    $formAttendance = $wb.Worksheets.Add(); $formAttendance.Name = "نموذج حضور"
    Title $formAttendance 8 "نموذج حضور دورة تدريبية"
    $formRows = @("اسم النشاط:", "كود النشاط:", "التاريخ:", "المدرب:", "القسم:")
    for ($i=0; $i -lt $formRows.Count; $i++) { Cell $formAttendance ($i+3) 1 $formRows[$i]; Cell $formAttendance ($i+3) 2 "" }
    $fh = @("م","كود الموظف","اسم الموظف","القسم","حضر؟","التوقيع","التقييم","ملاحظات")
    for ($c=0; $c -lt $fh.Count; $c++) { Cell $formAttendance 10 ($c+1) $fh[$c] }; Header $formAttendance 10 $fh.Count
    for ($r=11; $r -le 40; $r++) { Cell $formAttendance $r 1 ($r-10) }
    $formAttendance.PageSetup.Orientation = 2; $formAttendance.PageSetup.FitToPagesWide = 1; $formAttendance.PageSetup.FitToPagesTall = 1
    Finish $formAttendance

    $formEval = $wb.Worksheets.Add(); $formEval.Name = "نموذج تقييم"
    Title $formEval 6 "نموذج تقييم دورة تدريبية"
    $evalRows = @("اسم النشاط:", "كود النشاط:", "اسم الموظف:", "كود الموظف:", "التاريخ:", "اسم المقيّم:")
    for ($i=0; $i -lt $evalRows.Count; $i++) { Cell $formEval ($i+3) 1 $evalRows[$i]; Cell $formEval ($i+3) 2 "" }
    Cell $formEval 11 1 "البند"; Cell $formEval 11 2 "الدرجة"; Cell $formEval 11 3 "الملاحظات"; Header $formEval 11 3
    $items = @("فهم المحتوى","تطبيق المهارة","الالتزام بإجراءات السلامة","التواصل والعمل الجماعي","النتيجة النهائية")
    for ($i=0; $i -lt $items.Count; $i++) { Cell $formEval ($i+12) 1 $items[$i] }
    $formEval.PageSetup.Orientation = 1; $formEval.PageSetup.FitToPagesWide = 1; $formEval.PageSetup.FitToPagesTall = 1
    Finish $formEval

    foreach ($s in $wb.Worksheets) { $s.Activate(); $s.Application.ActiveWindow.SplitRow = 3; $s.Application.ActiveWindow.FreezePanes = $true }
    $settings.Activate()
    $wb.SaveAs($output, 51)
    Write-Output $output
}
finally {
    if ($wb) { $wb.Close($true) }
    if ($excel) { $excel.Quit() }
    [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
}
