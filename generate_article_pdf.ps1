# generate_article_pdf.ps1
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = "Stop"

$outPdfPath = "C:\Users\user\Downloads\work11\골목엔_나홀로_사장_방안엔_쉬는_청년_고용단절_리포트.pdf"
$w = 1654
$h = 2338

# Color Palette
$cDark = [System.Drawing.Color]::FromArgb(24, 33, 47)         # Slate 900
$cNavy = [System.Drawing.Color]::FromArgb(30, 58, 138)        # Blue 900
$cBlue = [System.Drawing.Color]::FromArgb(37, 99, 235)        # Blue 600
$cLightBlue = [System.Drawing.Color]::FromArgb(239, 246, 255) # Blue 50
$cRed = [System.Drawing.Color]::FromArgb(220, 38, 38)         # Red 600
$cLightRed = [System.Drawing.Color]::FromArgb(254, 242, 242)  # Red 50
$cGray = [System.Drawing.Color]::FromArgb(100, 116, 139)      # Slate 500
$cLightGray = [System.Drawing.Color]::FromArgb(241, 245, 249) # Slate 100
$cBorder = [System.Drawing.Color]::FromArgb(203, 213, 225)    # Slate 300
$cWhite = [System.Drawing.Color]::White

function RectF([float]$x, [float]$y, [float]$w, [float]$h) {
    return New-Object System.Drawing.RectangleF($x, $y, $w, $h)
}

function Fill-RoundRect($g, $brush, [float]$x, [float]$y, [float]$width, [float]$height, [float]$radius) {
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = $radius * 2
    $path.AddArc($x, $y, $d, $d, 180, 90)
    $path.AddArc(($x + $width - $d), $y, $d, $d, 270, 90)
    $path.AddArc(($x + $width - $d), ($y + $height - $d), $d, $d, 0, 90)
    $path.AddArc($x, ($y + $height - $d), $d, $d, 90, 90)
    $path.CloseFigure()
    $g.FillPath($brush, $path)
    $path.Dispose()
}

function Draw-RoundRect($g, $pen, [float]$x, [float]$y, [float]$width, [float]$height, [float]$radius) {
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = $radius * 2
    $path.AddArc($x, $y, $d, $d, 180, 90)
    $path.AddArc(($x + $width - $d), $y, $d, $d, 270, 90)
    $path.AddArc(($x + $width - $d), ($y + $height - $d), $d, $d, 0, 90)
    $path.AddArc($x, ($y + $height - $d), $d, $d, 90, 90)
    $path.CloseFigure()
    $g.DrawPath($pen, $path)
    $path.Dispose()
}

# ==============================================================
# PAGE 1 RENDERING
# ==============================================================
Write-Output "Rendering Page 1..."
$bmp1 = New-Object System.Drawing.Bitmap($w, $h)
$g1 = [System.Drawing.Graphics]::FromImage($bmp1)
$g1.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$g1.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
$g1.Clear($cWhite)

# 1. Top Decorative Bar
$brushNavy = New-Object System.Drawing.SolidBrush($cNavy)
$g1.FillRectangle($brushNavy, 0, 0, $w, 16)

# 2. Header Category Badge & Date
$brushBlue = New-Object System.Drawing.SolidBrush($cBlue)
Fill-RoundRect $g1 $brushBlue 100 60 280 40 8
$fBadge = New-Object System.Drawing.Font("맑은 고딕", 12, [System.Drawing.FontStyle]::Bold)
$brushWhite = New-Object System.Drawing.SolidBrush($cWhite)
$sfCenter = New-Object System.Drawing.StringFormat
$sfCenter.Alignment = [System.Drawing.StringAlignment]::Center
$sfCenter.LineAlignment = [System.Drawing.StringAlignment]::Center
$g1.DrawString("기획 보도 · 노동시장 심층 리포트", $fBadge, $brushWhite, (RectF 100 60 280 40), $sfCenter)

$fMeta = New-Object System.Drawing.Font("맑은 고딕", 11, [System.Drawing.FontStyle]::Regular)
$brushGray = New-Object System.Drawing.SolidBrush($cGray)
$sfRight = New-Object System.Drawing.StringFormat
$sfRight.Alignment = [System.Drawing.StringAlignment]::Far
$g1.DrawString("2026. 11. 19 | 국가데이터처 8월 부가조사 심층 분석", $fMeta, $brushGray, (RectF 700 70 854 30), $sfRight)

# 3. Main Headline
$fTitle = New-Object System.Drawing.Font("맑은 고딕", 27, [System.Drawing.FontStyle]::Bold)
$brushDark = New-Object System.Drawing.SolidBrush($cDark)
$titleText = "골목엔 '나홀로 사장', 방 안엔 '쉬는 청년'… 무너진 고용 사다리"
$g1.DrawString($titleText, $fTitle, $brushDark, (RectF 100 120 1454 70))

# 4. Subheadings
$fSub = New-Object System.Drawing.Font("맑은 고딕", 13, [System.Drawing.FontStyle]::Regular)
$brushNavyLight = New-Object System.Drawing.SolidBrush($cNavy)
$subText = "■ 자영업자 '직원 둘 여력 없다'… '나홀로 사장' 고착화에 첫 일자리·알바 자리도 실종`n■ 청년층 '원하는 일자리 없어'… 1년 내 취업 희망률 19.9%로 20%선마저 붕괴"
$g1.DrawString($subText, $fSub, $brushNavyLight, (RectF 100 195 1454 65))

# Divider Line
$penLine = New-Object System.Drawing.Pen($cBorder, 2)
$g1.DrawLine($penLine, 100, 275, 1554, 275)

# 5. Lead Paragraph Box
$brushLightBlue = New-Object System.Drawing.SolidBrush($cLightBlue)
$penBlueBorder = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(191, 219, 254), 1.5)
Fill-RoundRect $g1 $brushLightBlue 100 295 1454 125 12
Draw-RoundRect $g1 $penBlueBorder 100 295 1454 125 12

$brushBlueBar = New-Object System.Drawing.SolidBrush($cBlue)
$g1.FillRectangle($brushBlueBar, 100, 305, 8, 105)

$fLead = New-Object System.Drawing.Font("맑은 고딕", 12.5, [System.Drawing.FontStyle]::Bold)
$leadText = "[세종=경제부] 골목상권의 자영업자들은 벼랑 끝에 몰려 직원을 내보내고 있고, 사회 진출을 앞둔 청년들은 눈높이에 맞는 일자리가 없어 아예 취업 의욕을 접고 있다. 민간 고용의 진입 사다리 역할을 해왔던 자영업의 일자리 흡수력이 바닥을 치면서, 청년층이 노동시장 진입 단계부터 이탈해 '쉬었음'이나 학업 연장으로 내몰리는 고용 단절 악순환이 고착화되고 있다."
$g1.DrawString($leadText, $fLead, $brushDark, (RectF 130 310 1400 100))

# 6. Section 1: 자영업 위기
$fSec = New-Object System.Drawing.Font("맑은 고딕", 18, [System.Drawing.FontStyle]::Bold)
$g1.DrawString("1. '알바 쓸 돈도 없다'… 고용 포기한 '나홀로 사장님'", $fSec, $brushNavy, (RectF 100 445 1454 40))

$fBody = New-Object System.Drawing.Font("맑은 고딕", 12, [System.Drawing.FontStyle]::Regular)
$sec1Text = "국가데이터처의 2026년 8월 비임금근로 부가조사 결과에 따르면, 국내 자영업 생태계는 매출 부진과 인건비·임대료 등 고정비 부담으로 인해 '고용원 없는 자영업자(나홀로 사장)' 중심 구조가 한층 심화됐다. 직원이나 아르바이트생을 고용해 가게를 운영하던 자영업자들이 한계를 버티지 못하고 1인 자영업으로 전환하거나, 사업 전환 및 폐업을 심각하게 고민하는 비율이 높게 유지되고 있다.`n`n자영업자들의 주당 평균 근로시간은 일반 임금근로자보다 훨씬 긴 장시간 노동에 시달리고 있지만 수익성은 악화되면서 신규 인력을 채용할 여력 자체가 완전히 사라진 상태다. 과거 방학이나 학업 기간 동안 청년들이 사회 경험을 쌓고 생활비를 벌던 '단기 일자리 디딤돌' 자체가 골목상권에서부터 사라진 셈이다."
$g1.DrawString($sec1Text, $fBody, $brushDark, (RectF 100 495 1454 150))

# 7. INFOGRAPHIC 1: 자영업·소상공인 고용 위기 지표 (2개 카드)
$cardY = 660.0
$cardW = 710.0
$cardH = 340.0

# Card 1 Left: 고용원 없는 자영업자
Fill-RoundRect $g1 (New-Object System.Drawing.SolidBrush($cLightGray)) 100 $cardY $cardW $cardH 12
Draw-RoundRect $g1 (New-Object System.Drawing.Pen($cBorder, 1.5)) 100 $cardY $cardW $cardH 12

Fill-RoundRect $g1 $brushNavy 120 ($cardY + 20) 240 34 6
$g1.DrawString("자영업 고용구조", (New-Object System.Drawing.Font("맑은 고딕", 11, [System.Drawing.FontStyle]::Bold)), $brushWhite, (RectF 120 ($cardY + 20) 240 34), $sfCenter)

$fCardTitle = New-Object System.Drawing.Font("맑은 고딕", 15, [System.Drawing.FontStyle]::Bold)
$g1.DrawString("'나홀로 사장' 비중 고착화", $fCardTitle, $brushDark, 120, ($cardY + 68))

# Big Stat Display
$fBigNum = New-Object System.Drawing.Font("맑은 고딕", 36, [System.Drawing.FontStyle]::Bold)
$brushRed = New-Object System.Drawing.SolidBrush($cRed)
$g1.DrawString("64.8%", $fBigNum, $brushRed, 120, ($cardY + 110))

$fStatDesc = New-Object System.Drawing.Font("맑은 고딕", 11.5, [System.Drawing.FontStyle]::Regular)
$g1.DrawString("전체 자영업자 10명 중 6.5명이 1인 자영업자", $fStatDesc, $brushGray, 120, ($cardY + 175))

# Visual Proportion Bar
$barY = $cardY + 215
$g1.FillRectangle((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(226, 232, 240))), 120, $barY, 670, 28)
$g1.FillRectangle($brushRed, 120, $barY, (670 * 0.648), 28)
$g1.DrawString("고용원 없음(나홀로) 64.8%", (New-Object System.Drawing.Font("맑은 고딕", 10, [System.Drawing.FontStyle]::Bold)), $brushWhite, 130, ($barY + 5))
$g1.DrawString("고용원 있음 35.2%", (New-Object System.Drawing.Font("맑은 고딕", 10, [System.Drawing.FontStyle]::Bold)), $brushDark, 570, ($barY + 5))

$card1Footer = "※ 매출 부진 및 인건비·임대료 부담으로 아르바이트생 고용 엄두 못 내`n※ 폐업 고려 자영업자 중 70% 이상이 '생계형 한계 자영업자'"
$g1.DrawString($card1Footer, (New-Object System.Drawing.Font("맑은 고딕", 10, [System.Drawing.FontStyle]::Regular)), $brushGray, (RectF 120 ($cardY + 260) 670 60))

# Card 2 Right: 자영업자 근로시간 vs 임금근로자
$card2X = 844.0
Fill-RoundRect $g1 (New-Object System.Drawing.SolidBrush($cLightGray)) $card2X $cardY $cardW $cardH 12
Draw-RoundRect $g1 (New-Object System.Drawing.Pen($cBorder, 1.5)) $card2X $cardY $cardW $cardH 12

Fill-RoundRect $g1 $brushBlue $card2X ($cardY + 20) 240 34 6
$g1.DrawString("근로시간 격차", (New-Object System.Drawing.Font("맑은 고딕", 11, [System.Drawing.FontStyle]::Bold)), $brushWhite, (RectF $card2X ($cardY + 20) 240 34), $sfCenter)

$g1.DrawString("장시간 노동 vs 수익성 악화", $fCardTitle, $brushDark, ($card2X + 20), ($cardY + 68))

$g1.DrawString("44.2시간", $fBigNum, $brushBlue, ($card2X + 20), ($cardY + 110))
$g1.DrawString("자영업자 주당 평균 취업시간 (임금근로자 36.2시간 대비 +8.0시간)", $fStatDesc, $brushGray, ($card2X + 20), ($cardY + 175))

# Comparison Bars
$cBarY = $cardY + 215
$g1.FillRectangle((New-Object System.Drawing.SolidBrush($cBlue)), ($card2X + 20), $cBarY, (670 * (44.2/50.0)), 22)
$g1.DrawString("자영업자: 44.2시간/주", (New-Object System.Drawing.Font("맑은 고딕", 9.5, [System.Drawing.FontStyle]::Bold)), $brushWhite, ($card2X + 28), ($cBarY + 3))

$g1.FillRectangle((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(148, 163, 184))), ($card2X + 20), ($cBarY + 28), (670 * (36.2/50.0)), 22)
$g1.DrawString("임금근로자: 36.2시간/주", (New-Object System.Drawing.Font("맑은 고딕", 9.5, [System.Drawing.FontStyle]::Bold)), $brushWhite, ($card2X + 28), ($cBarY + 31))

$card2Footer = "※ 주 50시간 이상 초과근무 자영업자 비율 28.5% 육박`n※ 더 오래 일하지만 손에 쥐는 실질 소득은 오히려 감소하는 추세"
$g1.DrawString($card2Footer, (New-Object System.Drawing.Font("맑은 고딕", 10, [System.Drawing.FontStyle]::Regular)), $brushGray, (RectF ($card2X + 20) ($cardY + 275) 670 50))

# 8. Section 2: 청년 취업 희망률 붕괴
$sec2Y = 1030.0
$g1.DrawString("2. '갈 만한 곳 없다'… 취업 희망률 19.9%, 20%선 붕괴", $fSec, $brushNavy, (RectF 100 $sec2Y 1454 40))

$sec2Text = "자영업의 고용 한파는 고스란히 청년층의 구직 절벽으로 이어졌다. 국가데이터처의 비경제활동인구 부가조사에 따르면, 일도 구직활동도 하지 않고 '그냥 쉰' 인구는 255만 1,000명으로 집계됐다. 겉보기에는 전년 대비 8만 9,000명 줄어든 것처럼 보이지만, 내막을 들여다보면 사정은 결코 밝지 않다.`n`n청년층(15~29세)과 30대 비경제활동인구 중 쉬는 이유로 '원하는 일자리를 찾기 어려워서'를 꼽은 비율은 각각 35.2%, 32.3%로 전 연령대 중 압도적 1위를 기록했다. 대기업·공채 축소로 양질의 정규직 일자리는 바늘구멍이 된 반면, 눈높이를 낮춰 첫발을 뗄 수 있는 소상공인·중소기업 일자리조차 씨가 마르자 아예 구직 시장을 떠나 학원(재학·수강)이나 자격증 준비로 숨어든 것이다.`n`n특히 비경제활동인구(1,629만 7,000명) 중 향후 1년 이내에 '취업이나 창업을 희망한다'고 답한 비율은 19.9%(324만 명)에 그쳐, 2016년 관련 통계 작성 이후 역대 최저 수준으로 떨어졌다. 청년 10명 중 8명은 당장 일할 생각조차 품지 못하는 구조적 무기력에 빠진 것이다."
$g1.DrawString($sec2Text, $fBody, $brushDark, (RectF 100 ($sec2Y + 50) 1454 250))

# Page 1 Footer
$g1.DrawLine($penLine, 100, 2240, 1554, 2240)
$fFooter = New-Object System.Drawing.Font("맑은 고딕", 10, [System.Drawing.FontStyle]::Regular)
$g1.DrawString("출처: 국가데이터처 2026년 8월 경제활동인구조사 부가조사(비임금근로, 비경제활동, 근로형태)", $fFooter, $brushGray, 100, 2260)
$g1.DrawString("1 / 2", $fFooter, $brushGray, (RectF 700 2260 854 30), $sfRight)

$page1Jpg = "C:\Users\user\Downloads\work11\page1.jpg"
$bmp1.Save($page1Jpg, [System.Drawing.Imaging.ImageFormat]::Jpeg)
$g1.Dispose()
$bmp1.Dispose()
Write-Output "Page 1 rendered."

# ==============================================================
# PAGE 2 RENDERING
# ==============================================================
Write-Output "Rendering Page 2..."
$bmp2 = New-Object System.Drawing.Bitmap($w, $h)
$g2 = [System.Drawing.Graphics]::FromImage($bmp2)
$g2.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$g2.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
$g2.Clear($cWhite)

# 1. Top Decorative Bar
$g2.FillRectangle($brushNavy, 0, 0, $w, 16)

# Top mini header
$g2.DrawString("심층 리포트 | 고용 단절과 노동시장 양극화 분석", $fBadge, $brushNavy, 100, 50)
$g2.DrawString("2026년 11월 특집", $fMeta, $brushGray, (RectF 700 52 854 30), $sfRight)
$g2.DrawLine($penLine, 100, 90, 1554, 90)

# 2. INFOGRAPHIC 2: 청년층 비경제활동 & 일자리 미스매치 시각화
$g2.DrawString("■ [인포그래픽] 청년층 구직 단념과 노동시장 미스매치 실태", $fSec, $brushNavy, 100, 115)

# Chart Container Box
$boxY = 170.0
$boxH = 430.0
Fill-RoundRect $g2 (New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(248, 250, 252))) 100 $boxY 1454 $boxH 12
Draw-RoundRect $g2 (New-Object System.Drawing.Pen($cBorder, 1.5)) 100 $boxY 1454 $boxH 12

# Sub-Section Left: 청년·30대 쉬는 이유 Bar Chart
$g2.DrawString("연령대별 '쉬었음' 주된 사유 1위 비교", $fCardTitle, $brushDark, 140, ($boxY + 25))

# Bars
$barStartY = $boxY + 75.0
# Bar 1: 15~29세
$g2.DrawString("청년층 (15~29세)", (New-Object System.Drawing.Font("맑은 고딕", 12, [System.Drawing.FontStyle]::Bold)), $brushDark, 140, $barStartY)
$g2.FillRectangle((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(226, 232, 240))), 140, ($barStartY + 30), 650, 32)
$g2.FillRectangle($brushRed, 140, ($barStartY + 30), (650 * (35.2/50.0)), 32)
$g2.DrawString("원하는 일자리를 찾기 어려워서: 35.2%", (New-Object System.Drawing.Font("맑은 고딕", 11, [System.Drawing.FontStyle]::Bold)), $brushWhite, 150, ($barStartY + 36))

# Bar 2: 30대
$bar2Y = $barStartY + 80.0
$g2.DrawString("30대", (New-Object System.Drawing.Font("맑은 고딕", 12, [System.Drawing.FontStyle]::Bold)), $brushDark, 140, $bar2Y)
$g2.FillRectangle((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(226, 232, 240))), 140, ($bar2Y + 30), 650, 32)
$g2.FillRectangle((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(239, 68, 68))), 140, ($bar2Y + 30), (650 * (32.3/50.0)), 32)
$g2.DrawString("원하는 일자리를 찾기 어려워서: 32.3%", (New-Object System.Drawing.Font("맑은 고딕", 11, [System.Drawing.FontStyle]::Bold)), $brushWhite, 150, ($bar2Y + 36))

# Bar 3: 전 연령 평균
$bar3Y = $bar2Y + 80.0
$g2.DrawString("전체 연령 평균 (1위: 몸이 좋지 않아서 34.4%)", (New-Object System.Drawing.Font("맑은 고딕", 12, [System.Drawing.FontStyle]::Bold)), $brushDark, 140, $bar3Y)
$g2.FillRectangle((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(226, 232, 240))), 140, ($bar3Y + 30), 650, 32)
$g2.FillRectangle((New-Object System.Drawing.SolidBrush($cNavy)), 140, ($bar3Y + 30), (650 * (18.9/50.0)), 32)
$g2.DrawString("원하는 일자리 없음: 18.9%", (New-Object System.Drawing.Font("맑은 고딕", 11, [System.Drawing.FontStyle]::Bold)), $brushWhite, 150, ($bar3Y + 36))

$g2.DrawString("※ 청년층·30대는 '건강 문제'보다 '양질의 일자리 부족(미스매치)'이 압도적 사유", (New-Object System.Drawing.Font("맑은 고딕", 10.5, [System.Drawing.FontStyle]::Regular)), $brushGray, 140, ($bar3Y + 75))

# Sub-Section Right: 취업 희망률 게이지 / KPI 카드
$kpiX = 860.0
$kpiW = 650.0
Fill-RoundRect $g2 (New-Object System.Drawing.SolidBrush($cWhite)) $kpiX ($boxY + 25) ($kpiW - 40) ($boxH - 50) 10
Draw-RoundRect $g2 (New-Object System.Drawing.Pen($cBorder, 1.2)) $kpiX ($boxY + 25) ($kpiW - 40) ($boxH - 50) 10

$g2.DrawString("비경제활동인구 1년 내 취업·창업 희망률", $fCardTitle, $brushNavy, ($kpiX + 25), ($boxY + 50))
$g2.DrawString("19.9%", (New-Object System.Drawing.Font("맑은 고딕", 44, [System.Drawing.FontStyle]::Bold)), $brushRed, ($kpiX + 25), ($boxY + 95))
$g2.DrawString("2016년 관련 통계 작성 이래 역대 3번째 최저치 (20%선 붕괴)", $fStatDesc, $brushDark, ($kpiX + 25), ($boxY + 175))

# Gauge Bar
$g2.FillRectangle((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(226, 232, 240))), ($kpiX + 25), ($boxY + 215), 550, 32)
$g2.FillRectangle($brushRed, ($kpiX + 25), ($boxY + 215), (550 * 0.199), 32)
$g2.DrawString("취업 희망 19.9% (324만명)", (New-Object System.Drawing.Font("맑은 고딕", 10, [System.Drawing.FontStyle]::Bold)), $brushWhite, ($kpiX + 32), ($boxY + 222))
$g2.DrawString("취업 비희망 80.1% (1,305만명)", (New-Object System.Drawing.Font("맑은 고딕", 10, [System.Drawing.FontStyle]::Bold)), $brushDark, ($kpiX + 280), ($boxY + 222))

$kpiDetail = "■ 비경제활동인구 1,629만 7,000명 중 10명 중 8명이 '구직 포기'`n■ 일할 능력과 의욕이 있는 청년층마저 '취업 포기'로 내몰리는 실정"
$g2.DrawString($kpiDetail, (New-Object System.Drawing.Font("맑은 고딕", 11, [System.Drawing.FontStyle]::Regular)), $brushGray, (RectF ($kpiX + 25) ($boxY + 270) 550 80))

# 3. Comprehensive Data Table
$tblY = 630.0
$g2.DrawString("■ [종합 분석표] 통계로 본 대한민국 일자리 단절과 양극화 구조", $fSec, $brushNavy, 100, $tblY)

$tableStartY = $tblY + 45.0
$thH = 45.0
$colWs = @(220.0, 280.0, 480.0, 474.0)

# Table Header
$curX = 100.0
$headers = @("조사 부문", "핵심 대상", "주요 통계 지표 및 수치", "노동시장 구조적 시사점")
for ($i = 0; $i -lt 4; $i++) {
    $wCol = $colWs[$i]
    $g2.FillRectangle($brushNavy, $curX, $tableStartY, $wCol, $thH)
    $g2.DrawRectangle((New-Object System.Drawing.Pen($cWhite, 1)), $curX, $tableStartY, $wCol, $thH)
    $g2.DrawString($headers[$i], (New-Object System.Drawing.Font("맑은 고딕", 11, [System.Drawing.FontStyle]::Bold)), $brushWhite, (RectF $curX $tableStartY $wCol $thH), $sfCenter)
    $curX += $wCol
}

# Table Rows Data
$tableRows = @(
    @{
        Sec = "비임금근로`n(11월 부가조사)";
        Target = "자영업자`n(약 662만 명)";
        Stat = "• 고용원 없는 자영업자 비중: 64.8%`n• 자영업자 주당 근로: 44.2시간 (+8.0시간)`n• 주된 애로: 매출 부진 및 인건비·임대료 부담";
        Impact = "소상공인 고용 여력 완전 상실`n→ 청년층 첫 아르바이트·단기 일자리 소멸"
    },
    @{
        Sec = "비경제활동`n(11월 부가조사)";
        Target = "청년·쉬었음 인구`n(1,629만 7천 명)";
        Stat = "• 1년 내 취업·창업 희망률: 19.9% (역대 최저치)`n• 15~29세 쉬는 이유 1위: 원하는 일자리 없음 (35.2%)`n• 30대 쉬는 이유 1위: 원하는 일자리 없음 (32.3%)";
        Impact = "대기업·공채 축소로 양질의 일자리 부족`n→ 취업 자체를 포기하며 학업·학원으로 도피"
    },
    @{
        Sec = "근로형태별`n(10월 부가조사)";
        Target = "임금근로자`n(2,252만 6천 명)";
        Stat = "• 청년층(29세 이하) 비정규직: 148만 2천명 (-6.4만명)`n• 정규직 월평균 임금: 405만 2천원`n• 비정규직 월평균 임금: 211만 2천원 (격차 194만원)";
        Impact = "취업 관문을 뚫어도 처우 양극화 심각`n→ '불안정한 비정규직 하느니 쉰다'는 인식 확산"
    }
)

$rY = $tableStartY + $thH
$rowH = 110.0
$penTblLine = New-Object System.Drawing.Pen($cBorder, 1)

foreach ($r in $tableRows) {
    $curX = 100.0
    $bgBrush = $(if ($r -eq $tableRows[1]) { New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(254, 242, 242)) } else { New-Object System.Drawing.SolidBrush($cWhite) })
    
    # Fill cells
    for ($i = 0; $i -lt 4; $i++) {
        $wCol = $colWs[$i]
        $g2.FillRectangle($bgBrush, $curX, $rY, $wCol, $rowH)
        $g2.DrawRectangle($penTblLine, $curX, $rY, $wCol, $rowH)
        $curX += $wCol
    }
    
    # Texts
    $fCellBold = New-Object System.Drawing.Font("맑은 고딕", 11, [System.Drawing.FontStyle]::Bold)
    $fCellText = New-Object System.Drawing.Font("맑은 고딕", 10.5, [System.Drawing.FontStyle]::Regular)
    
    $g2.DrawString($r.Sec, $fCellBold, $brushNavy, (RectF 100 $rY $colWs[0] $rowH), $sfCenter)
    $g2.DrawString($r.Target, $fCellBold, $brushDark, (RectF 320 $rY $colWs[1] $rowH), $sfCenter)
    $g2.DrawString($r.Stat, $fCellText, $brushDark, (RectF 610 ($rY + 12) ($colWs[2] - 20) ($rowH - 24)))
    $impactBrush = $(if ($r -eq $tableRows[1]) { $brushRed } else { $brushDark })
    $g2.DrawString($r.Impact, $fCellBold, $impactBrush, (RectF 1090 ($rY + 18) ($colWs[3] - 20) ($rowH - 24)))
    
    $rY += $rowH
}

# 4. Section 3: 종합 진단 및 정책적 제언
$sec3Y = $rY + 35.0
$g2.DrawString("3. 첫 일자리 디딤돌 실종… 노동시장 이중구조와 양극화의 덫", $fSec, $brushNavy, (RectF 100 $sec3Y 1454 40))

$sec3Text = "결국 10월과 11월에 걸쳐 발표된 국가데이터처의 부가조사 통계들은 '골목상권의 고용 붕괴'와 '청년 구직 단념'이 별개의 현상이 아니라 긴밀하게 맞물린 하나의 거대한 고용 단절 위기임을 실증한다.`n`n자영업자가 비용 부담을 견디지 못해 직원을 줄이면서 사회초년생들이 진입할 수 있는 1차 완충지대가 사라졌고, 대기업·공채 축소로 안정적인 정규직 취업문은 극도로 좁아졌다. 어렵게 취업 관문을 뚫더라도 정규직(월 405만 원)과 비정규직(월 211만 원)의 심각한 임금 격차를 지켜본 청년들은 '불안정한 일자리를 전전하느니 차라리 스펙을 쌓으며 쉬겠다'며 노동시장에서 발길을 돌리고 있다.`n`n노동시장 전문가들은 '자영업자의 고용 여력 상실과 청년층의 구직 단념은 동전의 양면'이라며 '골목상권 소상공인의 고용 유지 부담을 덜어주는 현실적 지원책과 함께, 청년들이 일 경험을 쌓고 정규 일자리로 안착할 수 있도록 징검다리형 고용 인프라를 대폭 확충해야 한다'고 강조했다."
$g2.DrawString($sec3Text, $fBody, $brushDark, (RectF 100 ($sec3Y + 50) 1454 250))

# Page 2 Footer
$g2.DrawLine($penLine, 100, 2240, 1554, 2240)
$g2.DrawString("출처: 국가데이터처 2026년 8월 경제활동인구조사 부가조사(비임금근로, 비경제활동, 근로형태)", $fFooter, $brushGray, 100, 2260)
$g2.DrawString("2 / 2", $fFooter, $brushGray, (RectF 700 2260 854 30), $sfRight)

$page2Jpg = "C:\Users\user\Downloads\work11\page2.jpg"
$bmp2.Save($page2Jpg, [System.Drawing.Imaging.ImageFormat]::Jpeg)
$g2.Dispose()
$bmp2.Dispose()
Write-Output "Page 2 rendered."

# ==============================================================
# COMBINE TO MULTI-PAGE PDF
# ==============================================================
Write-Output "Building 2-page PDF..."

$jpg1Bytes = [System.IO.File]::ReadAllBytes($page1Jpg)
$jpg2Bytes = [System.IO.File]::ReadAllBytes($page2Jpg)

$fs = [System.IO.File]::Create($outPdfPath)
$wPdf = New-Object System.IO.BinaryWriter($fs)

function W-Str([string]$s) {
    $b = [System.Text.Encoding]::ASCII.GetBytes($s)
    $wPdf.Write($b)
}

$offsets = New-Object System.Collections.Generic.List[long]

W-Str("%PDF-1.4`n")

# Obj 1: Catalog
$offsets.Add($fs.Position)
W-Str("1 0 obj`n<< /Type /Catalog /Pages 2 0 R >>`nendobj`n")

# Obj 2: Pages
$offsets.Add($fs.Position)
W-Str("2 0 obj`n<< /Type /Pages /Kids [ 3 0 R 4 0 R ] /Count 2 >>`nendobj`n")

# Obj 3: Page 1
$offsets.Add($fs.Position)
W-Str("3 0 obj`n<< /Type /Page /Parent 2 0 R /MediaBox [ 0 0 595.28 841.89 ] /Resources << /XObject << /Im1 5 0 R >> >> /Contents 7 0 R >>`nendobj`n")

# Obj 4: Page 2
$offsets.Add($fs.Position)
W-Str("4 0 obj`n<< /Type /Page /Parent 2 0 R /MediaBox [ 0 0 595.28 841.89 ] /Resources << /XObject << /Im2 6 0 R >> >> /Contents 8 0 R >>`nendobj`n")

# Obj 5: Image 1
$offsets.Add($fs.Position)
W-Str("5 0 obj`n<< /Type /XObject /Subtype /Image /Width $w /Height $h /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length $($jpg1Bytes.Length) >>`nstream`n")
$wPdf.Write($jpg1Bytes)
W-Str("`nendstream`nendobj`n")

# Obj 6: Image 2
$offsets.Add($fs.Position)
W-Str("6 0 obj`n<< /Type /XObject /Subtype /Image /Width $w /Height $h /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length $($jpg2Bytes.Length) >>`nstream`n")
$wPdf.Write($jpg2Bytes)
W-Str("`nendstream`nendobj`n")

# Content streams
$content1 = "q`n595.28 0 0 841.89 0 0 cm`n/Im1 Do`nQ`n"
$content2 = "q`n595.28 0 0 841.89 0 0 cm`n/Im2 Do`nQ`n"

# Obj 7: Contents 1
$offsets.Add($fs.Position)
W-Str("7 0 obj`n<< /Length $($content1.Length) >>`nstream`n$content1`nendstream`nendobj`n")

# Obj 8: Contents 2
$offsets.Add($fs.Position)
W-Str("8 0 obj`n<< /Length $($content2.Length) >>`nstream`n$content2`nendstream`nendobj`n")

# xref
$xrefPos = $fs.Position
W-Str("xref`n0 9`n0000000000 65535 f `n")
for ($i = 0; $i -lt $offsets.Count; $i++) {
    $posStr = $offsets[$i].ToString("D10")
    W-Str("$posStr 00000 n `n")
}

W-Str("trailer`n<< /Size 9 /Root 1 0 R >>`nstartxref`n$xrefPos`n%%EOF`n")

$wPdf.Close()
$fs.Close()

Remove-Item -Force $page1Jpg, $page2Jpg -ErrorAction SilentlyContinue

Write-Output "Successfully generated 2-page PDF: $outPdfPath (Size: $((Get-Item $outPdfPath).Length) bytes)"
