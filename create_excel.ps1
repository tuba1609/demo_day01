$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Add()
while ($wb.Sheets.Count -gt 1) { $wb.Sheets.Item($wb.Sheets.Count).Delete() }

function StyleHeader($cell, $text, $bgColor, $fontSize=12) {
    $cell.Value2 = $text
    $cell.Font.Bold = $true
    $cell.Font.Size = $fontSize
    $cell.Font.Color = 0xFFFFFF
    $cell.Interior.Color = $bgColor
    $cell.HorizontalAlignment = -4108
    $cell.VerticalAlignment = -4108
    $cell.WrapText = $true
}
function AddBorder($range) {
    $range.Borders.LineStyle = 1
    $range.Borders.Weight = 2
}

# ==============================
# SHEET 1: Tong Quan
# ==============================
$ws1 = $wb.Sheets.Item(1)
$ws1.Name = "1. Tong Quan"
$ws1.Tab.Color = 0xFF6B9D
$ws1.Columns("A").ColumnWidth = 30
$ws1.Columns("B").ColumnWidth = 50

$ws1.Range("A1:B1").Merge()
StyleHeader $ws1.Cells(1,1) "KE HOACH KENH YOUTUBE - KE CHUYEN DAY KY NANG SONG CHO BE" 0xC0504D 16
$ws1.Rows("1").RowHeight = 50

$ws1.Range("A3:B3").Merge()
$ws1.Cells(3,1).Value2 = "THONG TIN KENH"
$ws1.Cells(3,1).Font.Bold = $true; $ws1.Cells(3,1).Font.Size = 12
$ws1.Range("A3:B3").Interior.Color = 0xF2DCDB

$info = @(
    @("Ten kenh (goi y)", "Be Vui Hoc / Chuyen Ke Cho Be / Vuon Chuyen Than Ky"),
    @("Slogan", "Moi cau chuyen - Mot bai hoc cuoc song"),
    @("Linh vuc / Niche", "Ke chuyen thieu nhi + Day ky nang song cho be 3-10 tuoi"),
    @("Dinh dang video", "Hoat hinh 2D / Ke chuyen co tranh minh hoa / Nguoi that dong"),
    @("Doi tuong chinh", "Ba me, bo - co con nho tu 3-10 tuoi"),
    @("Doi tuong phu", "Giao vien mam non, tieu hoc / Ong ba"),
    @("Ngon ngu", "Tieng Viet (giong ke chuyen nhe nhang, than thien)"),
    @("Tan suat dang video", "2-3 video/tuan (ngay T3, T5, T7)"),
    @("Thoi luong video", "5-12 phut / video (phu hop be xem cung bo me)"),
    @("Ngay khoi dong", "[DD/MM/YYYY]"),
    @("Nguoi phu trach", "[Ten / Team]")
)
$r = 4
foreach ($row in $info) {
    $ws1.Cells($r,1).Value2 = $row[0]; $ws1.Cells($r,1).Font.Bold = $true
    $ws1.Cells($r,1).Interior.Color = 0xFCE4D6
    $ws1.Cells($r,2).Value2 = $row[1]
    $r++
}

$ws1.Range("A16:B16").Merge()
$ws1.Cells(16,1).Value2 = "MUC TIEU DU AN"
$ws1.Cells(16,1).Font.Bold = $true; $ws1.Cells(16,1).Font.Size = 12
$ws1.Range("A16:B16").Interior.Color = 0xE2EFDA

$goals = @(
    @("Muc tieu 3 thang", "500 subscribers | 30.000 luot xem | 10 video chat luong"),
    @("Muc tieu 6 thang", "3.000 subscribers | 150.000 luot xem | Ra mat series dau"),
    @("Muc tieu 12 thang", "Dat 1.000 sub + 4.000h xem (du dieu kien monetize)"),
    @("Muc tieu dai han", "Tro thanh kenh ke chuyen cho be hang dau Viet Nam"),
    @("Gia tri cot loi", "Noi dung lanh manh - Co giao duc - An toan cho tre em")
)
$r = 17
foreach ($row in $goals) {
    $ws1.Cells($r,1).Value2 = $row[0]; $ws1.Cells($r,1).Font.Bold = $true
    $ws1.Cells($r,1).Interior.Color = 0xEBF1DE
    $ws1.Cells($r,2).Value2 = $row[1]
    $r++
}

$ws1.Range("A23:B23").Merge()
$ws1.Cells(23,1).Value2 = "CAC SERIES KE HOACH"
$ws1.Cells(23,1).Font.Bold = $true; $ws1.Cells(23,1).Font.Size = 12
$ws1.Range("A23:B23").Interior.Color = 0xDCE6F1

$series = @(
    @("Series 1: Ky nang ban be", "Chuyen ve chia se, lam ban, giai quyet mau thuan - 6 tap"),
    @("Series 2: Ky nang gia dinh", "Yeu thuong ong ba, giup do bo me, biet on - 6 tap"),
    @("Series 3: Ky nang truong hoc", "Mo dung gio, trung thuc, hoc bai cham chi - 6 tap"),
    @("Series 4: Ky nang an toan", "Khong noi chuyen nguoi la, an toan duong pho - 4 tap"),
    @("Series 5: Ky nang cam xuc", "Biet bieu lo cam xuc, kiem soat gian du, dong cam - 6 tap"),
    @("Video don le", "Chuyen co tich Viet Nam, le tet, su kien dac biet")
)
$r = 24
foreach ($row in $series) {
    $ws1.Cells($r,1).Value2 = $row[0]; $ws1.Cells($r,1).Font.Bold = $true
    $ws1.Cells($r,1).Interior.Color = 0xDAE8FC
    $ws1.Cells($r,2).Value2 = $row[1]
    $r++
}

AddBorder $ws1.Range("A1:B30")

# ==============================
# SHEET 2: Ke Hoach Noi Dung
# ==============================
$ws2 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets.Item($wb.Sheets.Count))
$ws2.Name = "2. Ke Hoach Video"
$ws2.Tab.Color = 0xFF9900

$colW = @(5, 38, 20, 25, 12, 13, 13, 22, 14)
$h2 = @("STT","Tieu De Video / Cau Chuyen","Series / Chu De","Ky Nang Song Day","Thoi Luong","Ngay Quay","Ngay Dang","Ghi Chu","Trang Thai")
for ($c=1;$c -le 9;$c++) {
    StyleHeader $ws2.Cells(1,$c) $h2[$c-1] 0xC0504D 11
    $ws2.Columns($c).ColumnWidth = $colW[$c-1]
}
$ws2.Rows("1").RowHeight = 38

$vids = @(
    @("1","Gioi thieu kenh - Chuyen gi se co trong kenh nay?","Gioi thieu","Tao su tay mo, ket noi","5-7 phut","","Tuan 1","Video dau, can thu vi an tuong","Len y tuong"),
    @("2","Chuyen: Chiec banh sinh nhat cua Miu","Series 1: Ban be","Chia se, quan tam ban","7-9 phut","","Tuan 1","Ky nang chia se do vat","Len y tuong"),
    @("3","Chuyen: Khi Ban Tot Bi Om","Series 1: Ban be","Cham soc, dong cam","6-8 phut","","Tuan 2","Day be biet quan tam khi ban benh","Len y tuong"),
    @("4","Chuyen: Lam Lanh Sau Cai Cai","Series 1: Ban be","Giai quyet mau thuan","8-10 phut","","Tuan 2","Ky nang xin loi va tha thu","Len y tuong"),
    @("5","Chuyen: Giup Do Ong Ba","Series 2: Gia dinh","Hieu thao, biet on","7-9 phut","","Tuan 3","Ky nang yeu thuong ong ba","Len y tuong"),
    @("6","Chuyen: Bo Me Di Lam Ve Muon","Series 2: Gia dinh","Kien nhan, tin tuong","6-8 phut","","Tuan 3","Ky nang cho doi va tin tuong","Len y tuong"),
    @("7","Chuyen: Em Be Moi Sinh","Series 2: Gia dinh","Nhuong nhin, yeu thuong","8-10 phut","","Tuan 4","Cho be co anh/chi","Len y tuong"),
    @("8","Chuyen: Khong Bao Gio Di Hoc Muon","Series 3: Truong hoc","Ngan nap, dung gio","7-9 phut","","Tuan 4","Ky nang quan ly thoi gian","Len y tuong"),
    @("9","Chuyen: Bai Kiem Tra Bi Sai","Series 3: Truong hoc","Trung thuc, nhan loi","8-10 phut","","Tuan 5","Ky nang trung thuc","Len y tuong"),
    @("10","Chuyen: Ban Moi Den Lop","Series 3: Truong hoc","Hoa nhap, than thien","7-9 phut","","Tuan 5","Ky nang ket ban moi","Len y tuong"),
    @("11","Chuyen: Nguoi La Cho Ke Kem","Series 4: An toan","An toan voi nguoi la","8-10 phut","","Tuan 6","Ky nang an toan ca nhan","Len y tuong"),
    @("12","Chuyen: Qua Duong An Toan","Series 4: An toan","An toan giao thong","6-8 phut","","Tuan 6","Ky nang qua duong","Len y tuong"),
    @("13","Chuyen: Khi Be Dang Gian","Series 5: Cam xuc","Kiem soat gian du","8-10 phut","","Tuan 7","Ky nang dieu chinh cam xuc","Len y tuong"),
    @("14","Chuyen: Buon Vi Bi Che","Series 5: Cam xuc","Dong cam, tu tin","7-9 phut","","Tuan 7","Ky nang vuot qua che bai","Len y tuong"),
    @("15","Chuyen: Mua He Cua Be (Tong ket)","Dac biet","Nhieu ky nang tong hop","10-12 phut","","Tuan 8","Video dac biet mua he","Len y tuong")
)
$statusColors = @{"Len y tuong"=0xFFF2CC;"Dang lam"=0xFFD966;"Hoan thanh"=0xA9D18E;"Da dang"=0x70AD47;"Dang sua"=0xF4B942}
$r = 2
foreach ($row in $vids) {
    for ($c=1;$c -le 9;$c++) {
        $ws2.Cells($r,$c).Value2 = $row[$c-1]
        $ws2.Cells($r,$c).WrapText = $true
        if ($r % 2 -eq 0) { $ws2.Cells($r,$c).Interior.Color = 0xFEE9DA }
    }
    $status = $row[8]
    if ($statusColors.ContainsKey($status)) { $ws2.Cells($r,9).Interior.Color = $statusColors[$status] }
    $ws2.Rows($r).RowHeight = 30
    $r++
}
AddBorder $ws2.Range("A1:I$r")

# ==============================
# SHEET 3: Lich San Xuat
# ==============================
$ws3 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets.Item($wb.Sheets.Count))
$ws3.Name = "3. Lich San Xuat"
$ws3.Tab.Color = 0x70AD47

$ws3.Range("A1:I1").Merge()
StyleHeader $ws3.Cells(1,1) "LICH SAN XUAT - THANG DAU TIEN (KENH KE CHUYEN CHO BE)" 0x375623 14
$ws3.Rows("1").RowHeight = 40

$ws3.Columns("A").ColumnWidth = 32
$ws3.Columns("I").ColumnWidth = 30
for ($c=2;$c -le 8;$c++) { $ws3.Columns($c).ColumnWidth = 9 }

$dH = @("Cong Viec","T2","T3","T4","T5","T6","T7","CN","Ghi Chu")
for ($c=1;$c -le 9;$c++) {
    $ws3.Cells(2,$c).Value2 = $dH[$c-1]
    $ws3.Cells(2,$c).Font.Bold = $true
    $ws3.Cells(2,$c).Interior.Color = 0x375623
    $ws3.Cells(2,$c).Font.Color = 0xFFFFFF
    $ws3.Cells(2,$c).HorizontalAlignment = -4108
}

$weeks = @(
    @{Wk="TUAN 1 - Len y tuong & Chuan bi kich ban"; BG=0xEBF1DE; Tasks=@(
        @("Nghen cuu chu de, xac dinh series","X","X","","","","","","Xem trend video thieu nhi"),
        @("Viet kich ban chuyen 1 (gioi thieu kenh)","","","X","X","","","","Van phong pham, Google Docs"),
        @("Viet kich ban chuyen 2 (Series 1/tap 1)","","","","","X","","",""),
        @("Thiet ke nhan vat, background","","","","","","X","X","Canva, Procreate hoac thue")
    )},
    @{Wk="TUAN 2 - San xuat video 1 (Gioi thieu kenh)"; BG=0xD9EAD3; Tasks=@(
        @("Lam hoat hinh / ve tranh minh hoa","X","X","X","","","","","CapCut, Animaker"),
        @("Thu am loi ke chuyen","","","","X","","","","Micro, phong yen tinh"),
        @("Ghep am thanh + hieu ung","","","","","X","","",""),
        @("Them nhac nen thieu nhi ban quyen","","","","","","X","","Epidemic Sound / YouTube Audio"),
        @("Thiet ke thumbnail be dep","","","","","","","X","Canva - chu to, mau sac tuoi")
    )},
    @{Wk="TUAN 3 - Dang video 1 + San xuat video 2"; BG=0xEBF1DE; Tasks=@(
        @("Review lan cuoi + chuan bi SEO","X","","","","","","",""),
        @("DANG VIDEO 1 (gio vang 8h sang T3)","","X","","","","","","Dang vao T3, T5 hoac T7"),
        @("Phan hoi comment, tuong tac phu huynh","X","","X","","X","","","Quan trong voi kenh thieu nhi"),
        @("San xuat video 2 (Chuyen: Chiec banh)","","","X","X","X","X","",""),
        @("Dang video 2 (T5 hoac T7)","","","","X","","X","","")
    )},
    @{Wk="TUAN 4 - Danh gia & Ke hoach thang 2"; BG=0xD9EAD3; Tasks=@(
        @("San xuat video 3","X","X","X","","","","",""),
        @("Phan tich Analytics tuan 1-3","","","","X","","","","Xem CTR, watch time, tuoi nguoi xem"),
        @("Dang video 3","","","","","X","","",""),
        @("Chia se len Facebook groups ba me","","","","","","X","","Groups nuoi day con, day ky nang"),
        @("Bao cao tong ket + ke hoach thang 2","","","","","","","X","Dieu chinh theo phan hoi")
    )}
)
$r = 3
foreach ($wk in $weeks) {
    $ws3.Range("A${r}:I${r}").Merge()
    $ws3.Cells($r,1).Value2 = $wk.Wk
    $ws3.Cells($r,1).Font.Bold = $true; $ws3.Cells($r,1).Font.Size = 11
    $ws3.Range("A${r}:I${r}").Interior.Color = $wk.BG
    $r++
    foreach ($task in $wk.Tasks) {
        for ($c=1;$c -le 9;$c++) {
            $ws3.Cells($r,$c).Value2 = $task[$c-1]
            $ws3.Cells($r,$c).HorizontalAlignment = if ($c -eq 1 -or $c -eq 9) { -4131 } else { -4108 }
            if ($task[$c-1] -eq "X") {
                $ws3.Cells($r,$c).Interior.Color = 0x92D050
                $ws3.Cells($r,$c).Font.Bold = $true
            }
        }
        $r++
    }
}
AddBorder $ws3.Range("A2:I$r")

# ==============================
# SHEET 4: Ngan Sach
# ==============================
$ws4 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets.Item($wb.Sheets.Count))
$ws4.Name = "4. Ngan Sach"
$ws4.Tab.Color = 0xFF0000

$ws4.Range("A1:F1").Merge()
StyleHeader $ws4.Cells(1,1) "NGAN SACH DU AN - KENH KE CHUYEN DAY KY NANG CHO BE" 0xC00000 14
$ws4.Rows("1").RowHeight = 40
$ws4.Columns("A").ColumnWidth = 30
$ws4.Columns("B").ColumnWidth = 26
$ws4.Columns("C").ColumnWidth = 16
$ws4.Columns("D").ColumnWidth = 12
$ws4.Columns("E").ColumnWidth = 18
$ws4.Columns("F").ColumnWidth = 30

$bH = @("Hang Muc","San Pham / Dich Vu","Don Gia (VND)","So Luong","Thanh Tien (VND)","Ghi Chu")
for ($c=1;$c -le 6;$c++) {
    $ws4.Cells(2,$c).Value2 = $bH[$c-1]
    $ws4.Cells(2,$c).Font.Bold = $true
    $ws4.Cells(2,$c).Interior.Color = 0xC00000
    $ws4.Cells(2,$c).Font.Color = 0xFFFFFF
    $ws4.Cells(2,$c).HorizontalAlignment = -4108
}

$budget = @(
    @{Cat="THIET BI THU AM & HINH ANH"; BG=0xFFE7E7; Items=@(
        @("Micro thu am giong ke chuyen","Micro USB Blue Yeti / Samson","800000","1","800000","Quan trong nhat - giong ke chuyen"),
        @("Phu kien chong on (pop filter)","Pop filter + arm stand","150000","1","150000",""),
        @("Tai nghe kiem am","Sony / Audio-Technica","300000","1","300000","De nghe va chinh am thanh"),
        @("Dien thoai / May tinh bang","San co / Cu mua them","0","1","0","Quay nguoi that dong (neu can)")
    )},
    @{Cat="PHAN MEM SAN XUAT VIDEO"; BG=0xFFEFEF; Items=@(
        @("Phan mem lam hoat hinh","Animaker / Vyond (goi co ban)","500000","1","500000","Lam hoat hinh cho be"),
        @("Phan mem dung phim","CapCut Pro / DaVinci Resolve","179000","1","179000",""),
        @("Thiet ke nhan vat & background","Canva Pro","179000","1","179000","Template, nhan vat vector"),
        @("Nhac nen thieu nhi ban quyen","Epidemic Sound","200000","1","200000","Tranh copyright"),
        @("Hieu ung am thanh (SFX)","Envato Elements","150000","1","150000","Tieng dong vui cho be")
    )},
    @{Cat="NOI DUNG & KICH BAN"; BG=0xFFE7E7; Items=@(
        @("Thue nguoi viet kich ban (neu can)","Freelancer","200000","4","800000","4 kich ban dau tien"),
        @("Thue nguoi doc loi / dien xuat","Nguoi doc chuyen cho be","300000","4","1200000","Giong dep, than thien"),
        @("Mua sach tham khao ky nang song","Sach day ky nang tre em","150000","3","450000","Lam nguon noi dung")
    )},
    @{Cat="MARKETING & PHAT TRIEN KENH"; BG=0xFFEFEF; Items=@(
        @("Quang cao Facebook Ads (nhom ba me)","Boost vao groups nuoi day con","500000","1","500000","Target: ba me co con 3-10 tuoi"),
        @("Thiet ke logo kenh / banner","Freelancer Canva","250000","1","250000","Logo de thuong, mau sac nhe"),
        @("Intro / Outro animation","Envato / tu lam","200000","1","200000","Nhan biet thuong hieu kenh"),
        @("Cham soc Fanpage Facebook","Dang clip ngan, chot bai","0","1","0","Minh hoa tu video YouTube")
    )},
    @{Cat="CHI PHI VAN HANH"; BG=0xFFE7E7; Items=@(
        @("Internet & dien","Hang thang","200000","3","600000","3 thang dau"),
        @("Luu tru dam may (Google Drive)","Google One 200GB","49000","3","147000",""),
        @("Khoa hoc YouTube / Video cho tre","Udemy / Kyna","300000","1","300000","Hoc SEO & content tre em"),
        @("Du phong (10%)","","0","0","657700","~10% tong chi phi")
    )}
)

$r = 3
$total = 0
foreach ($cat in $budget) {
    $ws4.Range("A${r}:F${r}").Merge()
    $ws4.Cells($r,1).Value2 = $cat.Cat
    $ws4.Cells($r,1).Font.Bold = $true; $ws4.Cells($r,1).Font.Size = 11
    $ws4.Range("A${r}:F${r}").Interior.Color = $cat.BG
    $r++
    foreach ($item in $cat.Items) {
        $ws4.Cells($r,1).Value2 = $item[0]
        $ws4.Cells($r,2).Value2 = $item[1]
        $ws4.Cells($r,3).Value2 = [int]$item[2]; $ws4.Cells($r,3).NumberFormat = "#,##0"
        $ws4.Cells($r,4).Value2 = [int]$item[3]
        $ws4.Cells($r,5).Value2 = [int]$item[4]; $ws4.Cells($r,5).NumberFormat = "#,##0"
        $ws4.Cells($r,6).Value2 = $item[5]
        $total += [int]$item[4]
        $r++
    }
}
$ws4.Range("A${r}:D${r}").Merge()
$ws4.Cells($r,1).Value2 = "TONG CHI PHI DU KIEN (3 THANG DAU)"
$ws4.Cells($r,1).Font.Bold = $true; $ws4.Cells($r,1).Font.Size = 13
$ws4.Range("A${r}:F${r}").Interior.Color = 0xC00000
$ws4.Range("A${r}:F${r}").Font.Color = 0xFFFFFF
$ws4.Cells($r,5).Value2 = [double]$total; $ws4.Cells($r,5).NumberFormat = "#,##0"
$ws4.Cells($r,5).Font.Bold = $true; $ws4.Cells($r,5).Font.Size = 13
$ws4.Cells($r,6).Value2 = "VND - co the giam neu tu lam nhieu"
AddBorder $ws4.Range("A2:F$r")

# ==============================
# SHEET 5: KPIs Theo Doi
# ==============================
$ws5 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets.Item($wb.Sheets.Count))
$ws5.Name = "5. KPIs Theo Doi"
$ws5.Tab.Color = 0x7030A0

$ws5.Range("A1:G1").Merge()
StyleHeader $ws5.Cells(1,1) "THEO DOI KPIs - KENH KE CHUYEN DAY KY NANG CHO BE" 0x7030A0 14
$ws5.Rows("1").RowHeight = 40
$ws5.Columns("A").ColumnWidth = 32
for ($c=2;$c -le 6;$c++) { $ws5.Columns($c).ColumnWidth = 14 }
$ws5.Columns("G").ColumnWidth = 28

$kH = @("Chi So (KPI)","Thang 1","Thang 2","Thang 3","Thang 6","Thang 12","Ghi Chu")
for ($c=1;$c -le 7;$c++) {
    $ws5.Cells(2,$c).Value2 = $kH[$c-1]
    $ws5.Cells(2,$c).Font.Bold = $true
    $ws5.Cells(2,$c).Interior.Color = 0x7030A0
    $ws5.Cells(2,$c).Font.Color = 0xFFFFFF
    $ws5.Cells(2,$c).HorizontalAlignment = -4108
}

$kpis = @(
    @{Cat="TANG TRUONG KENH"; BG=0xD9D2E9},
    @{Row=@("So Subscribers","100","400","1.200","6.000","25.000","Muc tieu tich luy")},
    @{Row=@("Tong Luot Xem","4.000","16.000","45.000","250.000","1.200.000","")},
    @{Row=@("Gio Xem (Watch Hours)","80","350","1.000","5.000","22.000","Can 4000h de monetize")},
    @{Row=@("So Video Da Dang","6","12","20","44","96","~2-3 video/tuan")},
    @{Cat="HIEU SUAT VIDEO"; BG=0xD9D2E9},
    @{Row=@("CTR (Click-Through Rate)","4%","5%","6%","7%","8%","Kenh thieu nhi CTR thuong cao")},
    @{Row=@("Avg View Duration (%)","45%","50%","55%","60%","65%","Muc tieu > 55% - be xem het")},
    @{Row=@("Avg View Duration (phut)","3 phut","4 phut","5 phut","6 phut","7 phut","")},
    @{Row=@("Luot Like / Video","20","60","150","500","2.000","")},
    @{Row=@("Luot Comment (phu huynh)","8","20","50","150","600","Comment phu huynh rat co gia tri")},
    @{Row=@("Luot Share / Video","5","20","60","200","800","Share la chi so quan trong nhat")},
    @{Cat="DOANH THU (sau khi monetize)"; BG=0xD9D2E9},
    @{Row=@("AdSense (VND / thang)","0","0","0","800.000","4.000.000","RPM kenh thieu nhi cao ~$3-5")},
    @{Row=@("Tai tro / Sponsor","0","0","0","1.500.000","8.000.000","Do choi, sach, app giao duc")},
    @{Row=@("Ban merchandise (sach/sticker)","0","0","0","0","3.000.000","")},
    @{Cat="MANG XA HOI & CONG DONG"; BG=0xD9D2E9},
    @{Row=@("Facebook Page / Group Followers","200","600","1.500","6.000","20.000","Target ba me co con nho")},
    @{Row=@("TikTok (clip ngan tu video)","100","400","1.000","8.000","30.000","Repurpose clip 30-60 giay")},
    @{Row=@("Zalo Group phu huynh","50","150","400","2.000","8.000","Nhom ho tro, chia se kinh nghiem")},
    @{Cat="CHAT LUONG NOI DUNG"; BG=0xD9D2E9},
    @{Row=@("Ty le phan hoi tich cuc (%)","80%","85%","90%","92%","95%","Comment khen ngoi / tong comment")},
    @{Row=@("So luong khieu nai noi dung","0","0","0","0","0","Muc tieu: 0 khieu nai")},
    @{Row=@("Video dat 1.000+ luot xem","0","1","3","8","20","Theo doi video viral")}
)
$r = 3
foreach ($kpi in $kpis) {
    if ($kpi.ContainsKey("Cat")) {
        $ws5.Range("A${r}:G${r}").Merge()
        $ws5.Cells($r,1).Value2 = $kpi.Cat
        $ws5.Cells($r,1).Font.Bold = $true; $ws5.Cells($r,1).HorizontalAlignment = -4108
        $ws5.Range("A${r}:G${r}").Interior.Color = $kpi.BG
        $r++
    } else {
        $row = $kpi.Row
        for ($c=1;$c -le 7;$c++) {
            $ws5.Cells($r,$c).Value2 = $row[$c-1]
            $ws5.Cells($r,$c).Interior.Color = if ($r%2 -eq 0) { 0xEAD1F5 } else { 0xF9F1FF }
            $ws5.Cells($r,$c).HorizontalAlignment = if ($c -eq 1 -or $c -eq 7) { -4131 } else { -4108 }
        }
        $r++
    }
}
AddBorder $ws5.Range("A2:G$r")

# ==============================
# SHEET 6: Quy Trinh San Xuat
# ==============================
$ws6 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets.Item($wb.Sheets.Count))
$ws6.Name = "6. Quy Trinh"
$ws6.Tab.Color = 0x00B0F0

$ws6.Range("A1:E1").Merge()
StyleHeader $ws6.Cells(1,1) "QUY TRINH SAN XUAT VIDEO KE CHUYEN CHO BE" 0x00599C 14
$ws6.Rows("1").RowHeight = 40
$ws6.Columns("A").ColumnWidth = 7
$ws6.Columns("B").ColumnWidth = 24
$ws6.Columns("C").ColumnWidth = 42
$ws6.Columns("D").ColumnWidth = 14
$ws6.Columns("E").ColumnWidth = 28

$pH = @("Buoc","Giai Doan","Cong Viec Chi Tiet","Thoi Gian","Cong Cu / Luu Y")
for ($c=1;$c -le 5;$c++) {
    $ws6.Cells(2,$c).Value2 = $pH[$c-1]
    $ws6.Cells(2,$c).Font.Bold = $true
    $ws6.Cells(2,$c).Interior.Color = 0x00599C
    $ws6.Cells(2,$c).Font.Color = 0xFFFFFF
    $ws6.Cells(2,$c).HorizontalAlignment = -4108
}

$procs = @(
    @("1","CHON CHU DE & KICH BAN","Xac dinh ky nang song can day (VD: chia se, trung thuc...)","1 ngay","Danh sach ky nang theo do tuoi"),
    @("","","Kiem tra xem chu de da co video chua (tranh trung lap)","1 gio","YouTube Search, TubeBuddy"),
    @("","","Viet outline: tinh huong - van de - giai quyet - bai hoc","2 gio","Google Docs"),
    @("","","Viet kich ban day du (doi thoai, cam xuc nhan vat)","3-4 gio","Giu ngon ngu don gian, de hieu"),
    @("","","Review kich ban (dam bao phu hop lua tuoi 3-10)","1 gio","Tranh noi dung phuc tap / so hai"),
    @("2","THIET KE NHAN VAT & CANH","Thiet ke / chon nhan vat chinh (ten, ngoai hinh, tinh cach)","2-3 gio","Canva, Procreate, Adobe Express"),
    @("","","Ve / chon background phu hop tung canh chuyen","2 gio","Mau sac tuoi sang, than thien"),
    @("","","Chuan bi props am thanh (SFX: tieng cuoi, nhac vui)","1 gio","Freesound.org, Epidemic Sound"),
    @("3","THU AM LOI KE CHUYEN","Chuan bi phong thu am yen tinh","30 phut","Phong kin, tranh tieng on"),
    @("","","Thu am loi ke (giong ke chuyen nhe nhang, tinh cam)","1-2 gio","Micro USB, Audacity"),
    @("","","Chinh sua am thanh: xu ly on, can bang am luong","1 gio","Audacity / Adobe Audition"),
    @("4","LAM HOAT HINH / DUNG VIDEO","Ghep nhan vat vao background theo tung canh","2-3 gio","Animaker, CapCut, Canva"),
    @("","","Dong bo loi ke chuyen voi hinh anh","1-2 gio",""),
    @("","","Them SFX va nhac nen (nhac nhe, vui tuoi)","1 gio","Nhac phai ban quyen!"),
    @("","","Them hieu ung chuyen canh (transition nhe nhang)","1 gio",""),
    @("","","Them subtitle (neu can) cho be theo doi","30 phut","Font chu lon, de doc"),
    @("","","Review toan bo video - kiem tra noi dung, am thanh","1 gio",""),
    @("5","UPLOAD & SEO","Thiet ke thumbnail: hinh nhan vat dep + chu ngan gon","1 gio","Canva - mau sac noi bat"),
    @("","","Viet tieu de: co tu khoa + thu hut phu huynh","30 phut","VD: Chuyen be Miu - Day chia se"),
    @("","","Viet mo ta day du: tom tat chuyen + ky nang day","30 phut","Them hashtag #ketcauyen #kynangsong"),
    @("","","Them Chapter / Timestamp trong mo ta","15 phut","Tang watch time, de navigate"),
    @("","","Them Cards, End Screen de xem them video","15 phut","Goi y video Series tiep theo"),
    @("","","Dat lich dang: T3/T5/T7 luc 8h sang","5 phut","Gio phu huynh cho be xem sang som"),
    @("6","QUANG BA & CONG DONG","Chia se clip ngan len TikTok, Facebook Reels","30 phut","Clip 30-60 giay, them CTA"),
    @("","","Dang vao groups Facebook: ba me, ky nang tre em","30 phut","Viet post than thien, khong spam"),
    @("","","Gui thong bao Zalo group phu huynh","10 phut",""),
    @("7","THEO DOI & PHAN HOI","Tra loi comment cua phu huynh (trong 24h)","30 phut","Rat quan trong de xay dung tin tuong"),
    @("","","Ghi nhan y kien: be thich / khong thich dieu gi","30 phut",""),
    @("","","Phan tich Analytics: CTR, watch time, tuoi nguoi xem","1 gio","YouTube Studio Analytics"),
    @("","","Ghi chu cai tien cho video tiep theo","30 phut","Google Docs - luu ho so du an")
)
$stepBG = @{"1"=0xFCE4D6;"2"=0xFFF2CC;"3"=0xE2EFDA;"4"=0xDDEBF7;"5"=0xEDEDED;"6"=0xE8D5F5;"7"=0xD9E1F2}
$curBG = 0xFFFFFF
$r = 3
foreach ($step in $procs) {
    if ($step[0] -ne "") { $curBG = $stepBG[$step[0]] }
    for ($c=1;$c -le 5;$c++) {
        $ws6.Cells($r,$c).Value2 = $step[$c-1]
        $ws6.Cells($r,$c).Interior.Color = $curBG
        $ws6.Cells($r,$c).WrapText = $true
        if ($step[0] -ne "" -and $c -le 2) { $ws6.Cells($r,$c).Font.Bold = $true }
    }
    $ws6.Rows($r).RowHeight = 22
    $r++
}
AddBorder $ws6.Range("A2:E$r")

# ==============================
# SHEET 7: Noi Dung Chi Tiet Series
# ==============================
$ws7 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets.Item($wb.Sheets.Count))
$ws7.Name = "7. Chi Tiet Series"
$ws7.Tab.Color = 0xFF6B9D

$ws7.Range("A1:F1").Merge()
StyleHeader $ws7.Cells(1,1) "CHI TIET CAC SERIES - KY NANG SONG CHO BE" 0xC0504D 14
$ws7.Rows("1").RowHeight = 40
$ws7.Columns("A").ColumnWidth = 22
$ws7.Columns("B").ColumnWidth = 8
$ws7.Columns("C").ColumnWidth = 32
$ws7.Columns("D").ColumnWidth = 28
$ws7.Columns("E").ColumnWidth = 22
$ws7.Columns("F").ColumnWidth = 20

$sH = @("Ten Series","Tap","Tieu De Chuyen","Ky Nang Day","Tinh Huong Chinh","Thong Diep")
for ($c=1;$c -le 6;$c++) {
    $ws7.Cells(2,$c).Value2 = $sH[$c-1]
    $ws7.Cells(2,$c).Font.Bold = $true
    $ws7.Cells(2,$c).Interior.Color = 0xC0504D
    $ws7.Cells(2,$c).Font.Color = 0xFFFFFF
    $ws7.Cells(2,$c).HorizontalAlignment = -4108
}

$seriesData = @(
    @{Name="Series 1: KY NANG BAN BE"; BG=0xFCE4D6; Eps=@(
        @("1","Chiec Banh Sinh Nhat Cua Miu","Chia se do vat","Miu co 1 chiec banh muon giu het","Chia se se lam ban be vui ve hon"),
        @("2","Khi Ban Tot Bi Om","Quan tam, dong cam","Bon bi om, Miu khong biet lam gi","Hoi tham va o ben ban khi ban buon"),
        @("3","Lam Lanh Sau Cai Cai","Xin loi va tha thu","Miu va Bon cai nhau vi do choi","Xin loi dung luc se giu duoc ban be"),
        @("4","Ban Moi Trong Lop","Ket ban voi nguoi la","Co be moi, ngo ngac, khong ai choi","Chu dong lam quen se co ban moi"),
        @("5","Khi Ban Noi Doi","Xu ly khi ban doi tra","Bon noi doi Miu va bi lo","Su that luon tot hon long doi"),
        @("6","Ban Be Khac Nhau Van Ok","Chap nhan su khac biet","Ban Miu bi khuyet tat chan","Moi nguoi deu co gia tri rieng")
    )},
    @{Name="Series 2: KY NANG GIA DINH"; BG=0xFFF2CC; Eps=@(
        @("1","Giup Do Ong Ba","Hieu thao, biet on","Miu ngai giup ong ba vi muon choi","Giup do ong ba la cach noi yeu thuong"),
        @("2","Bo Me Di Lam Ve Muon","Kien nhan, tin tuong","Miu so va khoc khi bo me ve muon","Bo me luon quay ve - hay tin tuong ho"),
        @("3","Em Be Moi Sinh","Nhuong nhin, yeu em","Miu ghen ty khi co em be","Anh chi lon con bao ve em nho"),
        @("4","Dem Khong Co Bo Me","Tu lap, khong so","Lan dau Miu ngu nha ong ba","Tu lam duoc nhieu viec la truong thanh"),
        @("5","Bua Com Gia Dinh","Tron trong, on hoa","Miu muon an nhanh de choi game","Bua com la thoi gian cua ca nha"),
        @("6","Gia Dinh Khong Hoan Hao","Chap nhan, yeu thuong","Ban Miu co gia dinh bo me ly hon","Moi gia dinh deu co yeu thuong rieng")
    )},
    @{Name="Series 3: KY NANG TRUONG HOC"; BG=0xE2EFDA; Eps=@(
        @("1","Khong Bao Gio Di Hoc Muon","Ngan nap, quan ly gio","Miu thuong ngu quen, di muon","Chuan bi truoc se khong bao gio muon"),
        @("2","Bai Kiem Tra Bi Sai","Trung thuc, nhan loi","Miu chep bai ban bi thay bat","Thua nhan sai lam con hon gian doi"),
        @("3","Ban Moi Den Lop","Giup do ban moi","Co ban moi ngo ngac, khong ai giup","Moi nguoi deu can duoc giup do luc moi"),
        @("4","Khong Muon Hoc Bai","Kien tri, co gang","Miu luoi hoc, muon xem TV","Hoc bai xong moi choi duoc lau hon"),
        @("5","Khi Bi Che Cuoi O Lop","Tu tin, ban linh","Miu doc sai bi ban cuoi","Sai lam la cach de hoc - dung so"),
        @("6","Du An Nhom","Hop tac, chia viec","Nhom cua Miu ai cung muon lam truong","Lam cung nhau se tot hon lam mot minh")
    )},
    @{Name="Series 4: KY NANG AN TOAN"; BG=0xDAE8FC; Eps=@(
        @("1","Nguoi La Cho Ke Kem","Canh giac nguoi la","Nguoi la cho Miu kem de du di","Khong nhan do nguoi la - du co dep den dau"),
        @("2","Qua Duong An Toan","An toan giao thong","Miu suyt bi xe khi qua duong","Den xanh, nhin ky, di bo voi nguoi lon"),
        @("3","Bi Lac Tren Pho","Xu ly khi bi lac","Miu bi lac o sieu thi","Dung khoc - tim bao ve, bao ten bo me"),
        @("4","Bao Mat Ca Nhan Tren Mang","An toan internet","Ban Miu chia se thong tin cho nguoi la","Khong chia se thong tin ca nhan online")
    )}
)
$r = 3
foreach ($s in $seriesData) {
    $ws7.Range("A${r}:F${r}").Merge()
    $ws7.Cells($r,1).Value2 = $s.Name
    $ws7.Cells($r,1).Font.Bold = $true; $ws7.Cells($r,1).Font.Size = 11
    $ws7.Range("A${r}:F${r}").Interior.Color = $s.BG
    $r++
    foreach ($ep in $s.Eps) {
        $ws7.Cells($r,1).Value2 = $s.Name.Split(":")[0].Trim()
        $ws7.Cells($r,2).Value2 = "Tap " + $ep[0]
        $ws7.Cells($r,3).Value2 = $ep[1]
        $ws7.Cells($r,4).Value2 = $ep[2]
        $ws7.Cells($r,5).Value2 = $ep[3]
        $ws7.Cells($r,6).Value2 = $ep[4]
        for ($c=1;$c -le 6;$c++) { $ws7.Cells($r,$c).WrapText = $true }
        if ($r%2 -eq 0) {
            for ($c=1;$c -le 6;$c++) { $ws7.Cells($r,$c).Interior.Color = 0xF5F5F5 }
        }
        $ws7.Rows($r).RowHeight = 30
        $r++
    }
}
AddBorder $ws7.Range("A2:F$r")

# ==============================
# SHEET 8: Cong Cu AI
# ==============================
$ws8 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets.Item($wb.Sheets.Count))
$ws8.Name = "8. Cong Cu AI"
$ws8.Tab.Color = 0x00B050

$ws8.Range("A1:G1").Merge()
StyleHeader $ws8.Cells(1,1) "CONG CU AI CHO KENH KE CHUYEN - DAY KY NANG SONG CHO BE" 0x1F6B36 15
$ws8.Rows("1").RowHeight = 48

$ws8.Columns("A").ColumnWidth = 8
$ws8.Columns("B").ColumnWidth = 20
$ws8.Columns("C").ColumnWidth = 22
$ws8.Columns("D").ColumnWidth = 28
$ws8.Columns("E").ColumnWidth = 16
$ws8.Columns("F").ColumnWidth = 14
$ws8.Columns("G").ColumnWidth = 26

$aH = @("Buoc","Giai Doan San Xuat","Cong Cu AI","Chuc Nang Cu The","Muc Gia","Muc Do Uu Tien","Luu Y")
for ($c=1;$c -le 7;$c++) {
    $ws8.Cells(2,$c).Value2 = $aH[$c-1]
    $ws8.Cells(2,$c).Font.Bold = $true
    $ws8.Cells(2,$c).Interior.Color = 0x1F6B36
    $ws8.Cells(2,$c).Font.Color = 0xFFFFFF
    $ws8.Cells(2,$c).HorizontalAlignment = -4108
    $ws8.Cells(2,$c).WrapText = $true
}
$ws8.Rows("2").RowHeight = 32

$aiTools = @(
    @{Cat="BUOC 1: VIET KICH BAN & Y TUONG"; BG=0xE2EFDA},
    @("1","Viet kich ban chuyen","ChatGPT / Claude","Prompt: 'Viet kich ban ke chuyen ve ky nang [X] cho be 5 tuoi, co nhan vat de thuong, tinh huong thuc te, thong diep ro rang'","Mien phi / $20/thang","CAO NHAT","Claude rat tot cho van ban Viet"),
    @("1","Brainstorm y tuong","Google Gemini","Goi y chuyen theo chu de ky nang song, phu hop vung mien VN","Mien phi","CAO","Hieu van hoa Viet tot"),
    @("1","SEO tieu de video","ChatGPT + TubeBuddy AI","Tim tu khoa phu huynh hay tim khi cho be xem: 'chuyen be...', 'day be...'","$9/thang (TB)","TRUNG BINH","Ket hop 2 cong cu"),
    @{Cat="BUOC 2: TAO HINH ANH & NHAN VAT"; BG=0xFFF2CC},
    @("2","Tao nhan vat chinh","Midjourney v6","Prompt: 'cute Vietnamese cartoon child character, simple, friendly, 2D flat design, Pixar style'","$10/thang","CAO NHAT","Nhat quan bo nhan vat"),
    @("2","Ve background / canh","DALL-E 3 (ChatGPT)","Tao canh: truong hoc, gia dinh, san choi Viet Nam phong cach cartoon","Mien phi (co han)","CAO","Mien phi trong ChatGPT Plus"),
    @("2","Chinh sua & ghep hinh","Adobe Firefly","Xoa nen, thay doi mau sac nhan vat, chuan hoa phong cach","Free / $10/thang","TRUNG BINH","Tich hop san trong Adobe"),
    @("2","Thiet ke nhanh","Canva AI (Magic Design)","Tao thumbnail, poster, banner kenh tu text prompt","Mien phi / $15/thang","CAO","De dung nhat cho nguoi moi"),
    @{Cat="BUOC 3: GIONG DOC & AM THANH"; BG=0xFCE4D6},
    @("3","Giong ke chuyen tieng Viet","FPT.AI Voice","Text-to-speech tieng Viet nhieu giong (Nam Bac, giong nu/nam, giong tre em)","Mien phi co han","CAO NHAT","Tieng Viet tu nhien nhat"),
    @("3","Giong doc chat luong cao","ElevenLabs","Clone giong hoac dung giong san, rat tu nhien, ho tro tieng Viet","$5-22/thang","CAO NHAT","Chat luong tot nhat hien tai"),
    @("3","Backup giong doc","Murf AI","Thu vien giong da dang, co tieng Viet co ban","$19/thang","TRUNG BINH","Goi y neu ElevenLabs dat"),
    @("3","Tao nhac nen thieu nhi","Suno AI","Prompt: 'happy children background music, ukulele, soft, Vietnamese style, no lyrics'","Mien phi / $8/thang","CAO","Nhac doc quyen, khong copyright"),
    @("3","Tao nhac nen 2","Udio","Tuong tu Suno, them tuy chon nhac cu, nhac dan toc","Mien phi beta","TRUNG BINH","Dang trong giai doan beta"),
    @{Cat="BUOC 4: LAM HOAT HINH & VIDEO"; BG=0xDAE8FC},
    @("4","Lam hoat hinh slide 2D","Animaker AI","Keo tha nhan vat, them dong tac, phong cach cartoon dep cho be","$10-20/thang","CAO NHAT","De dung, nhieu template thieu nhi"),
    @("4","Tao video tu hinh anh","Kling AI (Kuaishou)","Bien hinh anh nhan vat thanh video co chuyen dong tu nhien","Mien phi co han","CAO","Chuyen dong rat dep"),
    @("4","Chinh sua video AI","CapCut AI","Tu dong cat ghep, them phu de, hieu ung, chinh mau sac","Mien phi / Pro $10","CAO NHAT","De dung, nhieu template VN"),
    @("4","Hieu ung chuyen canh","Runway ML","Tao hieu ung bien doi canh dep, hieu ung dac biet","$12/thang","TRUNG BINH","Dung cho canh dac biet"),
    @("4","Avatar nguoi ke chuyen","HeyGen","Tao avatar AI ke chuyen that, dong bo moi voi giong doc","$24/thang","TRUNG BINH","Neu muon co nguoi that dong"),
    @{Cat="BUOC 5: PHAN DE & BIEN TAP"; BG=0xEDEDED},
    @("5","Tu dong tao phu de","Kapwing AI","Upload video, AI tu dong tao phu de tieng Viet chinh xac","Mien phi / $16/thang","CAO","Giup be theo doi dung hon"),
    @("5","Chinh sua phu de","Descript","Sua phu de nhu sua van ban, cat video bang cach xoa chu","$12/thang","CAO","Rat thuan tien cho bien tap"),
    @("5","Kiem tra noi dung an toan","YouTube Safety Check","Xem lai nguon: YouTube chay Content ID tu dong","Mien phi","BAT BUOC","Tranh copyright nhac, hinh"),
    @{Cat="BUOC 6: DANG VIDEO & SEO"; BG=0xE8D5F5},
    @("6","Toi uu SEO YouTube","VidIQ AI (Boost)","Phan tich tu khoa, goi y tieu de/tag, xem score SEO","Mien phi / $10/thang","CAO NHAT","Plugin Chrome de cai"),
    @("6","Toi uu SEO bo sung","TubeBuddy AI","A/B test thumbnail, tag explorer, best time to post","Mien phi / $9/thang","CAO","Ket hop voi VidIQ"),
    @("6","Tao thumbnail hap dan","Canva AI + Magic Write","Tao thumbnail: nhan vat chinh + chu ngan gon + mau sac noi bat","Mien phi","CAO NHAT","Font chu lon > 40pt"),
    @("6","Viet mo ta video","Claude / ChatGPT","Viet mo ta 200-500 chu, co tu khoa, goi y xem them video khac","Mien phi","CAO","Nen dung template co san"),
    @{Cat="BUOC 7: PHAN TICH & PHAT TRIEN"; BG=0xFCE4D6},
    @("7","Phan tich hieu suat","YouTube Analytics + AI","Xem CTR, watch time, tuoi nguoi xem theo doi tu dong","Mien phi","BAT BUOC","Xem moi tuan 1 lan"),
    @("7","Hieu content de tang view","Opus Clip AI","Tu dong cat clip ngan viral tu video dai, dang TikTok/Reels","$15/thang","CAO","Tang traffic tu TikTok ve YT"),
    @("7","Quan ly lich dang","Buffer AI Assist","Lap lich dang nhieu kenh, goi y gio vang, theo doi tuong tac","$6/thang","TRUNG BINH","Quan ly FB, TikTok, YT 1 cho"),
    @("7","Nghien cuu doi thu","ChatGPT + YouTube","Phan tich kenh ke chuyen lon (Peppa Pig VN, Thoi Gian Vang...)","Mien phi","CAO","Hoc tu kenh thanh cong")
)

$stepBG2 = @{"1"=0xE2EFDA;"2"=0xFFF2CC;"3"=0xFCE4D6;"4"=0xDAE8FC;"5"=0xF2F2F2;"6"=0xE8D5F5;"7"=0xFFF2CC}
$priorityColor = @{"CAO NHAT"=0x00B050;"CAO"=0x92D050;"TRUNG BINH"=0xFFD966;"BAT BUOC"=0xFF0000}
$r = 3
foreach ($item in $aiTools) {
    if ($item -is [hashtable]) {
        $ws8.Range("A${r}:G${r}").Merge()
        $ws8.Cells($r,1).Value2 = $item.Cat
        $ws8.Cells($r,1).Font.Bold = $true; $ws8.Cells($r,1).Font.Size = 11
        $ws8.Range("A${r}:G${r}").Interior.Color = $item.BG
        $ws8.Cells($r,1).HorizontalAlignment = -4131
        $ws8.Rows($r).RowHeight = 24
        $r++
    } else {
        for ($c=1;$c -le 7;$c++) {
            $ws8.Cells($r,$c).Value2 = $item[$c-1]
            $ws8.Cells($r,$c).WrapText = $true
            if ($r%2 -eq 0) { $ws8.Cells($r,$c).Interior.Color = 0xF7F7F7 }
        }
        $prio = $item[5]
        if ($priorityColor.ContainsKey($prio)) {
            $ws8.Cells($r,6).Interior.Color = $priorityColor[$prio]
            $ws8.Cells($r,6).Font.Bold = $true
            $ws8.Cells($r,6).Font.Color = if ($prio -eq "CAO NHAT" -or $prio -eq "BAT BUOC") { 0xFFFFFF } else { 0x000000 }
        }
        $ws8.Cells($r,6).HorizontalAlignment = -4108
        $ws8.Rows($r).RowHeight = 36
        $r++
    }
}
AddBorder $ws8.Range("A2:G$r")

# Row tong ket AI stack de xuat
$r++
$ws8.Range("A${r}:G${r}").Merge()
$ws8.Cells($r,1).Value2 = "AI STACK TIEU CHUAN DE XUAT (chi phi ~500K-1.5 trieu VND/thang): FPT.AI Voice + Midjourney + Animaker + CapCut AI + VidIQ + Suno AI + ChatGPT/Claude"
$ws8.Cells($r,1).Font.Bold = $true; $ws8.Cells($r,1).Font.Size = 11
$ws8.Range("A${r}:G${r}").Interior.Color = 0x1F6B36
$ws8.Range("A${r}:G${r}").Font.Color = 0xFFFFFF
$ws8.Cells($r,1).WrapText = $true
$ws8.Rows($r).RowHeight = 36
AddBorder $ws8.Range("A${r}:G${r}")

# ==============================
# Save
# ==============================
$path = "D:\KeHoach_Kenh_YouTube.xlsx"
$wb.SaveAs($path)
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host "DONE: $path"
