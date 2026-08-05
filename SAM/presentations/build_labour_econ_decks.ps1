$ErrorActionPreference = 'Stop'

$TemplatePath = 'C:\Users\OEM\Documents\GitHub\TVHENZ\e61 Projects\Presentations\e61 template.potx'
$OutDir = 'C:\Users\OEM\Documents\GitHub\COVID_LabourSupply\SAM\presentations'
$AssetDir = Join-Path $OutDir 'assets'
$PreviewDir = Join-Path $OutDir 'previews'
New-Item -ItemType Directory -Force -Path $OutDir, $PreviewDir | Out-Null

function RGB([int]$r, [int]$g, [int]$b) { return $r + 256 * $g + 65536 * $b }

$C = @{
    Teal       = RGB 0 79 84
    Aqua       = RGB 58 167 172
    PaleAqua   = RGB 198 245 244
    Orange     = RGB 232 111 61
    PaleOrange = RGB 255 239 231
    LightGrey  = RGB 240 243 243
    MidGrey    = RGB 166 166 166
    DarkGrey   = RGB 74 78 80
    Black      = RGB 0 0 0
    White      = RGB 255 255 255
    Red        = RGB 172 46 47
    PaleRed    = RGB 253 235 235
    Green      = RGB 41 119 79
    PaleGreen  = RGB 232 246 238
}

$Font = 'Proxima Nova Rg'
$Serif = 'GT Sectra Fine'
$MathFont = 'Cambria Math'

function Add-Text {
    param(
        $Slide, [double]$X, [double]$Y, [double]$W, [double]$H,
        [string]$Text, [double]$Size = 16, [int]$Color = $C.Black,
        [bool]$Bold = $false, [int]$Align = 1, [string]$FontName = $Font,
        [double]$Margin = 2
    )
    $shape = $Slide.Shapes.AddTextbox(1, $X, $Y, $W, $H)
    $shape.Line.Visible = 0
    $shape.Fill.Visible = 0
    $shape.TextFrame.WordWrap = -1
    $shape.TextFrame.MarginLeft = $Margin
    $shape.TextFrame.MarginRight = $Margin
    $shape.TextFrame.MarginTop = $Margin
    $shape.TextFrame.MarginBottom = $Margin
    $shape.TextFrame.TextRange.Text = $Text
    $shape.TextFrame.TextRange.Font.Name = $FontName
    $shape.TextFrame.TextRange.Font.Size = $Size
    $shape.TextFrame.TextRange.Font.Color.RGB = $Color
    $shape.TextFrame.TextRange.Font.Bold = $(if ($Bold) { -1 } else { 0 })
    $shape.TextFrame.TextRange.ParagraphFormat.Alignment = $Align
    $shape.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 4
    return $shape
}

function Add-Rect {
    param($Slide, [double]$X, [double]$Y, [double]$W, [double]$H,
          [int]$Fill, [int]$Line = $Fill, [double]$Radius = 1)
    $shapeType = $(if ($Radius -gt 0) { 5 } else { 1 })
    $shape = $Slide.Shapes.AddShape($shapeType, $X, $Y, $W, $H)
    $shape.Fill.ForeColor.RGB = $Fill
    $shape.Line.ForeColor.RGB = $Line
    $shape.Line.Weight = 0.75
    return $shape
}

function Add-Line {
    param($Slide, [double]$X1, [double]$Y1, [double]$X2, [double]$Y2,
          [int]$Color = $C.Teal, [double]$Weight = 1.5, [bool]$Arrow = $false)
    $shape = $Slide.Shapes.AddLine($X1, $Y1, $X2, $Y2)
    $shape.Line.ForeColor.RGB = $Color
    $shape.Line.Weight = $Weight
    if ($Arrow) { $shape.Line.EndArrowheadStyle = 3 }
    return $shape
}

function Add-PictureFit {
    param($Slide, [string]$Path, [double]$X, [double]$Y, [double]$W, [double]$H)
    if (-not (Test-Path -LiteralPath $Path)) { throw "Missing image: $Path" }
    return $Slide.Shapes.AddPicture($Path, 0, -1, $X, $Y, $W, $H)
}

function Add-Ticker {
    param($Slide, [string[]]$Labels, [int]$Active)
    $segment = 960.0 / $Labels.Count
    for ($i = 0; $i -lt $Labels.Count; $i++) {
        $fill = $(if ($i -eq $Active) { $C.Teal } else { $C.LightGrey })
        $textColor = $(if ($i -eq $Active) { $C.White } else { $C.MidGrey })
        Add-Rect $Slide ($i * $segment) 0 $segment 24 $fill $fill 0 | Out-Null
        Add-Text $Slide ($i * $segment) 2 $segment 18 $Labels[$i] 9.5 $textColor ($i -eq $Active) 2 | Out-Null
    }
}

function New-ContentSlide {
    param($Pres, [string]$Title, [string[]]$Ticker, [int]$Active, [string]$Kicker = '')
    $slide = $Pres.Slides.AddSlide($Pres.Slides.Count + 1, $Pres.SlideMaster.CustomLayouts.Item(22))
    Add-Ticker $slide $Ticker $Active
    $titleSize = $(if ($Title.Length -gt 72) { 21 } elseif ($Title.Length -gt 52) { 23 } else { 26 })
    Add-Text $slide 38 37 880 48 $Title $titleSize $C.Black $true 1 | Out-Null
    Add-Rect $slide 39 88 145 4 $C.Teal $C.Teal 0 | Out-Null
    if ($Kicker) { Add-Text $slide 745 69 165 18 $Kicker 9 $C.MidGrey $false 3 | Out-Null }
    return $slide
}

function Add-Source {
    param($Slide, [string]$Text)
    Add-Text $Slide 45 508 760 16 $Text 7.5 $C.MidGrey $false 1 | Out-Null
}

function Add-Callout {
    param($Slide, [string]$Text, [double]$X = 55, [double]$Y = 430,
          [double]$W = 840, [double]$H = 58, [string]$Kind = 'teal')
    switch ($Kind) {
        'orange' { $fill = $C.PaleOrange; $line = $C.Orange; $color = $C.Orange }
        'red'    { $fill = $C.PaleRed; $line = $C.Red; $color = $C.Red }
        'green'  { $fill = $C.PaleGreen; $line = $C.Green; $color = $C.Green }
        default  { $fill = $C.PaleAqua; $line = $C.Aqua; $color = $C.Teal }
    }
    Add-Rect $Slide $X $Y $W $H $fill $line 1 | Out-Null
    $size = $(if ($Text.Length -gt 170) { 12.5 } elseif ($Text.Length -gt 110) { 14 } else { 15.5 })
    Add-Text $Slide ($X + 14) ($Y + 8) ($W - 28) ($H - 16) $Text $size $color $true 1 | Out-Null
}

function Add-BulletSlide {
    param($Pres, [string]$Title, [string[]]$Ticker, [int]$Active,
          [string[]]$Bullets, [string]$Takeaway = '', [string]$TakeawayKind = 'teal',
          [string]$Source = '')
    $slide = New-ContentSlide $Pres $Title $Ticker $Active
    $text = ($Bullets | ForEach-Object { "• $_" }) -join "`r`n"
    $chars = $text.Length
    $size = $(if ($chars -gt 1000) { 13 } elseif ($chars -gt 760) { 14.5 } else { 16.5 })
    $height = $(if ($Takeaway) { 306 } else { 360 })
    Add-Text $slide 58 112 840 $height $text $size $C.DarkGrey $false 1 | Out-Null
    if ($Takeaway) { Add-Callout $slide $Takeaway 55 428 840 58 $TakeawayKind }
    if ($Source) { Add-Source $slide $Source }
    return $slide
}

function Add-Panel {
    param($Slide, [double]$X, [double]$Y, [double]$W, [double]$H,
          [string]$Heading, [string]$Body, [string]$Style = 'teal')
    switch ($Style) {
        'orange' { $head = $C.Orange; $bodyFill = $C.PaleOrange; $bodyColor = $C.DarkGrey }
        'red'    { $head = $C.Red; $bodyFill = $C.PaleRed; $bodyColor = $C.DarkGrey }
        'green'  { $head = $C.Green; $bodyFill = $C.PaleGreen; $bodyColor = $C.DarkGrey }
        'grey'   { $head = $C.DarkGrey; $bodyFill = $C.LightGrey; $bodyColor = $C.DarkGrey }
        default  { $head = $C.Teal; $bodyFill = $C.LightGrey; $bodyColor = $C.DarkGrey }
    }
    Add-Rect $Slide $X $Y $W 42 $head $head 0 | Out-Null
    Add-Text $Slide ($X + 10) ($Y + 8) ($W - 20) 27 $Heading 14.5 $C.White $true 1 | Out-Null
    Add-Rect $Slide $X ($Y + 42) $W ($H - 42) $bodyFill $bodyFill 0 | Out-Null
    $size = $(if ($Body.Length -gt 650) { 12 } elseif ($Body.Length -gt 470) { 13 } else { 14.5 })
    Add-Text $Slide ($X + 13) ($Y + 54) ($W - 26) ($H - 66) $Body $size $bodyColor $false 1 | Out-Null
}

function Add-TwoColumnSlide {
    param($Pres, [string]$Title, [string[]]$Ticker, [int]$Active,
          [string]$LeftHeading, [string]$LeftBody,
          [string]$RightHeading, [string]$RightBody,
          [string]$LeftStyle = 'teal', [string]$RightStyle = 'orange',
          [string]$Takeaway = '', [string]$Source = '')
    $slide = New-ContentSlide $Pres $Title $Ticker $Active
    $h = $(if ($Takeaway) { 295 } else { 360 })
    Add-Panel $slide 48 112 420 $h $LeftHeading $LeftBody $LeftStyle
    Add-Panel $slide 492 112 420 $h $RightHeading $RightBody $RightStyle
    if ($Takeaway) { Add-Callout $slide $Takeaway 55 431 840 55 'teal' }
    if ($Source) { Add-Source $slide $Source }
    return $slide
}

function Add-ThreeCardsSlide {
    param($Pres, [string]$Title, [string[]]$Ticker, [int]$Active,
          [string[]]$Heads, [string[]]$Bodies, [string]$Takeaway = '')
    $slide = New-ContentSlide $Pres $Title $Ticker $Active
    $xs = @(45, 349, 653)
    $styles = @('teal', 'orange', 'grey')
    for ($i = 0; $i -lt 3; $i++) { Add-Panel $slide $xs[$i] 125 266 270 $Heads[$i] $Bodies[$i] $styles[$i] }
    if ($Takeaway) { Add-Callout $slide $Takeaway 55 425 840 62 'teal' }
    return $slide
}

function Add-ImageSlide {
    param($Pres, [string]$Title, [string[]]$Ticker, [int]$Active,
          [string]$Image, [string]$Caption = '', [string]$Source = '',
          [double]$X = 70, [double]$Y = 108, [double]$W = 820, [double]$H = 365)
    $slide = New-ContentSlide $Pres $Title $Ticker $Active
    Add-PictureFit $slide $Image $X $Y $W $H | Out-Null
    if ($Caption) { Add-Text $slide 70 475 820 28 $Caption 10.5 $C.DarkGrey $false 1 | Out-Null }
    if ($Source) { Add-Source $slide $Source }
    return $slide
}

function Add-TableSlide {
    param($Pres, [string]$Title, [string[]]$Ticker, [int]$Active,
          [string[]]$Headers, [object[]]$Rows, [double[]]$Widths,
          [string]$Takeaway = '', [string]$Source = '')
    $slide = New-ContentSlide $Pres $Title $Ticker $Active
    $x0 = 48.0; $y0 = 116.0; $tableW = 864.0
    $sum = ($Widths | Measure-Object -Sum).Sum
    $scaled = $Widths | ForEach-Object { $tableW * $_ / $sum }
    $rowH = [Math]::Min(57, 330.0 / ($Rows.Count + 1))
    $x = $x0
    for ($col = 0; $col -lt $Headers.Count; $col++) {
        Add-Rect $slide $x $y0 $scaled[$col] $rowH $C.Teal $C.White 0 | Out-Null
        Add-Text $slide ($x + 5) ($y0 + 6) ($scaled[$col] - 10) ($rowH - 12) $Headers[$col] 12.5 $C.White $true 1 | Out-Null
        $x += $scaled[$col]
    }
    for ($r = 0; $r -lt $Rows.Count; $r++) {
        $x = $x0; $fill = $(if ($r % 2 -eq 0) { $C.LightGrey } else { $C.White })
        for ($col = 0; $col -lt $Headers.Count; $col++) {
            Add-Rect $slide $x ($y0 + ($r + 1) * $rowH) $scaled[$col] $rowH $fill $C.White 0 | Out-Null
            $txt = [string]$Rows[$r][$col]
            $size = $(if ($txt.Length -gt 95) { 9.5 } elseif ($txt.Length -gt 55) { 10.5 } else { 12 })
            Add-Text $slide ($x + 5) ($y0 + ($r + 1) * $rowH + 5) ($scaled[$col] - 10) ($rowH - 10) $txt $size $C.DarkGrey ($col -eq 0) 1 | Out-Null
            $x += $scaled[$col]
        }
    }
    if ($Takeaway) { Add-Callout $slide $Takeaway 55 432 840 55 'teal' }
    if ($Source) { Add-Source $slide $Source }
    return $slide
}

function Add-SectionSlide {
    param($Pres, [string]$Number, [string]$Title, [string]$Subtitle = '')
    $slide = $Pres.Slides.AddSlide($Pres.Slides.Count + 1, $Pres.SlideMaster.CustomLayouts.Item(26))
    Add-Text $slide 110 168 110 55 $Number 34 $C.Aqua $true 1 | Out-Null
    Add-Text $slide 110 225 750 90 $Title 36 $C.White $true 1 | Out-Null
    if ($Subtitle) { Add-Text $slide 112 320 710 60 $Subtitle 17 $C.White $false 1 $Serif | Out-Null }
    return $slide
}

function Add-TitleText {
    param($Slide, [string]$Title, [string]$Subtitle, [string]$Meta)
    Add-Text $Slide 62 108 690 145 $Title 40 $C.White $true 1 | Out-Null
    Add-Text $Slide 65 265 650 82 $Subtitle 21 $C.White $false 1 $Serif | Out-Null
    Add-Text $Slide 65 430 650 40 $Meta 12 $C.White $false 1 | Out-Null
}

function Add-FlowBox {
    param($Slide, [double]$X, [double]$Y, [double]$W, [double]$H,
          [string]$Heading, [string]$Body, [string]$Style = 'teal')
    switch ($Style) {
        'aqua' { $fill = $C.Aqua; $text = $C.White }
        'orange' { $fill = $C.Orange; $text = $C.White }
        'grey' { $fill = $C.LightGrey; $text = $C.DarkGrey }
        'red' { $fill = $C.PaleRed; $text = $C.Red }
        default { $fill = $C.Teal; $text = $C.White }
    }
    Add-Rect $Slide $X $Y $W $H $fill $fill 1 | Out-Null
    Add-Text $Slide ($X + 10) ($Y + 9) ($W - 20) 25 $Heading 14 $text $true 2 | Out-Null
    if ($Body) { Add-Text $Slide ($X + 10) ($Y + 36) ($W - 20) ($H - 43) $Body 11.5 $text $false 2 | Out-Null }
}

function Add-ModelStateSlide {
    param($Pres, [string[]]$Ticker)
    $slide = New-ContentSlide $Pres 'The model has two worker types and two labour-market states' $Ticker 1
    Add-FlowBox $slide 55 160 235 145 'Non-employed: Uⱼ' "Outside option bⱼ + ℓ`r`nChoose search sⱼ`r`nReceive/accept offers" 'aqua'
    Add-FlowBox $slide 665 160 235 145 'Employed: Wⱼ' "Receive wage w − hⱼ`r`nContinue or exit`r`nMatch surplus governs hazard" 'teal'
    Add-Line $slide 295 210 655 210 $C.Teal 2 $true | Out-Null
    Add-Text $slide 375 167 205 34 'find: fⱼ = aⱼ[1−exp(−csⱼ)]' 14 $C.Teal $true 2 $MathFont | Out-Null
    Add-Line $slide 655 270 295 270 $C.Orange 2 $true | Out-Null
    Add-Text $slide 372 277 215 34 'exit: δⱼ = δ₀ + ξ + AΛ(·)' 14 $C.Orange $true 2 $MathFont | Out-Null
    Add-Text $slide 190 342 580 60 'j ∈ {R, N}: R is benefit-eligible / Australian; N is ineligible / New Zealand comparison.' 17 $C.DarkGrey $false 2 | Out-Null
    Add-Callout $slide 'This is a partial-equilibrium interpretation model. The two types share job opportunities, but their outside options and residual work wedges differ.' 90 420 780 63 'teal'
    return $slide
}

function Add-CalibrationFlowSlide {
    param($Pres, [string[]]$Ticker)
    $slide = New-ContentSlide $Pres 'Calibration is sequential: each empirical cell is assigned a job' $Ticker 2
    $boxes = @(
        @('1', 'Pre exits', 'Fit baseline + amplitude of smooth exit hazard'),
        @('2', 'Pre JFRs', 'Fit matching efficiency μ and search-cost level κ'),
        @('3', 'NZ post exit', 'Fit common background destruction shock ξ'),
        @('4', 'AU post exit', 'Fit residual differential work-surplus wedge hᴿ'),
        @('5', 'NZ post JFR', 'Fit post matching efficiency μ; hold κ fixed'),
        @('6', 'AU post JFR', 'Held out as a prediction in the restricted model')
    )
    $xs = @(50, 345, 640, 50, 345, 640); $ys = @(130, 130, 130, 285, 285, 285)
    for ($i=0; $i -lt 6; $i++) {
        $style = $(if ($i -eq 5) { 'orange' } elseif ($i -ge 3) { 'aqua' } else { 'teal' })
        Add-FlowBox $slide $xs[$i] $ys[$i] 270 118 ($boxes[$i][0] + '. ' + $boxes[$i][1]) $boxes[$i][2] $style
        if ($i -lt 2) { Add-Line $slide ($xs[$i]+270) 189 $xs[$i+1] 189 $C.DarkGrey 1.5 $true | Out-Null }
        if ($i -ge 3 -and $i -lt 5) { Add-Line $slide ($xs[$i]+270) 344 $xs[$i+1] 344 $C.DarkGrey 1.5 $true | Out-Null }
    }
    Add-Callout $slide 'Teaching point: «held out» is useful only if specification choices were not tuned after seeing the held-out result.' 110 438 740 50 'orange'
    return $slide
}

function Add-ColliderSlide {
    param($Pres, [string[]]$Ticker)
    $slide = New-ContentSlide $Pres 'Conditioning on future non-employment creates a collider' $Ticker 2
    Add-FlowBox $slide 60 160 205 95 'Policy exposure' 'Australian eligibility / benefit receipt' 'teal'
    Add-FlowBox $slide 378 160 205 95 'Included in sample' '≥4 weeks non-employed later in 2020' 'red'
    Add-FlowBox $slide 695 160 205 95 'Latent fragility' 'Unobserved attachment, health, employer risk' 'grey'
    Add-Line $slide 265 208 368 208 $C.Red 2 $true | Out-Null
    Add-Line $slide 695 208 593 208 $C.Red 2 $true | Out-Null
    Add-FlowBox $slide 378 320 205 90 'Measured transition' 'Job finding or employment exit' 'orange'
    Add-Line $slide 480 255 480 312 $C.Red 2 $true | Out-Null
    Add-Line $slide 797 255 590 335 $C.DarkGrey 1.5 $true | Out-Null
    Add-Callout $slide 'Selecting on the middle box opens a non-causal path between policy exposure and latent fragility. The exit sample is especially mechanical because inclusion requires a later exit.' 82 430 796 60 'red'
    return $slide
}

function Add-RiskSetSlide {
    param($Pres, [string[]]$Ticker)
    $slide = New-ContentSlide $Pres 'The redesign starts with a pre-policy cohort and weekly risk sets' $Ticker 3
    Add-FlowBox $slide 50 140 225 110 'Freeze cohort at t₀' 'Residence + working age + 2019 labour-market attachment only' 'teal'
    Add-FlowBox $slide 365 140 225 110 'Build person-week states' 'Employed, not employed, missing/reporting, multiple jobs' 'aqua'
    Add-FlowBox $slide 680 140 225 110 'Define two hazards' 'Uₜ → Eₜ₊₁ and Eₜ → Uₜ₊₁ on contemporaneous denominators' 'teal'
    Add-Line $slide 275 195 355 195 $C.DarkGrey 1.6 $true | Out-Null
    Add-Line $slide 590 195 670 195 $C.DarkGrey 1.6 $true | Out-Null
    Add-Panel $slide 110 300 330 108 'Job-finding risk set' "Everyone not in recorded payroll employment in week t.`r`nAdd flexible duration dependence." 'green'
    Add-Panel $slide 520 300 330 108 'Employment-exit risk set' "Everyone in recorded payroll employment in week t.`r`nPersistence is an outcome definition—not inclusion." 'orange'
    Add-Callout $slide 'Run 1-, 2-, and 4-week persistence definitions; report weekly denominators, events and censoring by nationality.' 110 438 740 50 'teal'
    return $slide
}

function Add-PolicyTimelineSlide {
    param($Pres, [string[]]$Ticker)
    $slide = New-ContentSlide $Pres 'The «treatment date» is a sequence of announcements and implementation dates' $Ticker 0
    Add-Line $slide 85 260 875 260 $C.DarkGrey 2 $false | Out-Null
    $events = @(
        @('22 Mar', 'Supplement announced', 'orange'),
        @('24 Mar', 'Mutual obligations suspended', 'aqua'),
        @('30 Mar', 'JobKeeper announced', 'teal'),
        @('27 Apr', '$550 supplement begins', 'orange'),
        @('9 Jun', 'Modified obligations resume', 'aqua'),
        @('25 Sep', 'Supplement reduced', 'orange')
    )
    $xs = @(82, 220, 358, 496, 634, 772)
    for ($i=0; $i -lt 6; $i++) {
        Add-Line $slide ($xs[$i]+48) 240 ($xs[$i]+48) 280 $C.DarkGrey 1.5 $false | Out-Null
        $y = $(if ($i % 2 -eq 0) { 130 } else { 300 })
        Add-FlowBox $slide $xs[$i] $y 112 96 $events[$i][0] $events[$i][1] $events[$i][2]
        if ($i % 2 -eq 0) { Add-Line $slide ($xs[$i]+48) 226 ($xs[$i]+48) 240 $C.DarkGrey 1.2 $false | Out-Null }
        else { Add-Line $slide ($xs[$i]+48) 280 ($xs[$i]+48) 300 $C.DarkGrey 1.2 $false | Out-Null }
    }
    Add-Callout $slide 'Estimate announcement, entitlement and payment windows separately where feasible; otherwise label the estimand as exposure to the March JobSeeker policy bundle.' 90 430 780 58 'orange'
    Add-Source $slide 'Dates follow the manuscript and prior review; final deck should be updated from the projectʼs exact legal/administrative timeline.'
    return $slide
}

function Initialise-Presentation {
    param($Ppt, [string]$OutputPptx, [string]$Title, [string]$Subtitle)
    $pres = $Ppt.Presentations.Open($TemplatePath, $false, $true, $false)
    for ($i = $pres.Slides.Count; $i -ge 1; $i--) { if ($i -ne 2) { $pres.Slides.Item($i).Delete() } }
    $titleSlide = $pres.Slides.Item(1)
    Add-TitleText $titleSlide $Title $Subtitle 'Labour economics teaching deck | Based on manuscript v10 and structural model 15 | August 2026'
    return $pres
}

function Save-Presentation {
    param($Pres, [string]$PptxPath, [string]$PdfPath, [int[]]$PreviewSlides)
    $Pres.SaveAs($PptxPath, 24)
    $Pres.SaveAs($PdfPath, 32)
    $base = [IO.Path]::GetFileNameWithoutExtension($PptxPath)
    foreach ($n in $PreviewSlides) {
        if ($n -le $Pres.Slides.Count) {
            $path = Join-Path $PreviewDir ("{0}_slide_{1:D2}.png" -f $base, $n)
            $Pres.Slides.Item($n).Export($path, 'PNG', 1600, 900)
        }
    }
}

function Build-StructuralDeck {
    param($Ppt)
    $pptx = Join-Path $OutDir 'COVID_labour_supply_structural_model_teaching_deck.pptx'
    $pdf = Join-Path $OutDir 'COVID_labour_supply_structural_model_teaching_deck.pdf'
    $T = @('Motivation', 'Model anatomy', 'Calibration', 'Assessment', 'Roadmap')
    $pres = Initialise-Presentation $Ppt $pptx 'Structural interpretation of emergency unemployment assistance' 'Model 15: mechanisms, calibration, weaknesses and a submission-ready research agenda'

    Add-ThreeCardsSlide $pres 'What this deck is designed to teach' $T 0 `
        @('Read the mechanism', 'Interrogate identification', 'Use the model honestly') `
        @(
            "Translate benefits, work wedges and matching conditions into search, acceptance and employment exits.",
            "See which moments discipline which parameters—and why near-equivalent fits do not identify a unique mechanism.",
            "Separate a useful quantitative interpretation exercise from welfare analysis or a fully structural equilibrium model."
        ) 'The goal is not to defend every current modelling choice; it is to show what can be learned now and what must change before journal submission.' | Out-Null

    Add-ThreeCardsSlide $pres 'Bottom line on model 15' $T 0 `
        @('Real advance', 'Main empirical message', 'Binding limitation') `
        @(
            "Model 15 gives bounded, surplus-sensitive exit probabilities; shares a common surplus index across finding and exits; and genuinely holds out Australian post JFR.",
            "In the benchmark, the benefit-only job-finding DiD is −1.50 pp versus −1.31 pp in the full environment. For exits, nonlinear interaction is central.",
            "The selected benchmark is weakly identified. Near-fitting parameterisations imply held-out errors from roughly −2.84 to +0.48 pp and work wedges from 5% to 41% of wages."
        ) 'Treat the current model as a partially identified interpretation model. Do not present its point decomposition as uniquely estimated.' | Out-Null

    Add-TwoColumnSlide $pres 'Roadmap' $T 0 `
        'Part I — Learn the current model' "1. Empirical moments`r`n2. Worker problem and search`r`n3. Acceptance and exit hazards`r`n4. Firms, matching and dynamics`r`n5. Calibration mapping" `
        'Part II — Decide what to keep' "6. Benchmark results`r`n7. Identification audit`r`n8. Internal-consistency audit`r`n9. Minimal submission-ready model`r`n10. Ambitious extensions and paper write-up" `
        'teal' 'aqua' 'Slides deliberately alternate equations, intuition, diagnostics and recommended decisions.' | Out-Null

    Add-SectionSlide $pres '01' 'What must the model explain?' 'Start from transition moments; ask which mechanisms can rationalise them.' | Out-Null

    Add-ImageSlide $pres 'The draft supplies four transition cells—and two DiD contrasts' $T 0 (Join-Path $AssetDir 'empirical_did_cells.png') `
        'These numbers are provisional until the empirical risk sets and inference are rebuilt.' `
        'Source: manuscript Tables 3–4 and model-15 Targets().' 65 105 830 375 | Out-Null

    Add-TwoColumnSlide $pres 'Why add structure to a credible reduced-form design?' $T 0 `
        'Reduced form can estimate' "• The change in Australian-minus-NZ job finding`r`n• The change in Australian-minus-NZ payroll exits`r`n• Heterogeneity by predicted policy exposure`r`n• Dynamic responses around announcements" `
        'Structure can organise' "• Benefits versus common job-opportunity shocks`r`n• Search versus acceptance margins`r`n• Complementarity in match continuation`r`n• Counterfactual changes outside the observed episode" `
        'teal' 'orange' 'Structure does not repair a selected empirical sample. Revised target moments must come first.' | Out-Null

    Add-TableSlide $pres 'What changed from the manuscript model to version 15?' $T 0 `
        @('Dimension','Manuscript / earlier model','Model 15','Implication') `
        @(
            @('Employment exits','Largely exogenous or additively shifted','Smooth, bounded hazard of net flow surplus','Benefits and work wedges interact nonlinearly'),
            @('Post calibration','Fits both groupsʼ post JFRs','NZ post JFR fitted; AU post JFR held out','Creates a genuine validation target'),
            @('Search costs','May vary pre/post','Restricted model holds κ fixed','Common post shock is channelled through μ'),
            @('Decomposition','Narrative «health versus benefits»','Four scenario environments','Scenarios are not additive contributions'),
            @('Exit language','Match-quality interpretation','Static group-level logistic hazard','Must be described as reduced-form unless quality state is implemented')
        ) @(1.0,1.25,1.35,1.55) 'The paperʼs current structural section and abstract must be rewritten from model 15 outputs—not patched line by line.' `
        'Source: manuscript section 7; 15_structural_separation_model.jl.' | Out-Null

    Add-SectionSlide $pres '02' 'Model anatomy' 'Two types, two states, a common market, and a surplus-sensitive exit hazard.' | Out-Null

    Add-ModelStateSlide $pres $T | Out-Null

    Add-TwoColumnSlide $pres 'Flow payoffs create the net work-surplus index' $T 1 `
        'Eligible / Australian type R' "Outside: bᴿ + ℓ`r`nWork: w − hᴿ`r`nFlow surplus: zᴿ = w − hᴿ − bᴿ − ℓ`r`nThe supplement raises bᴿ and lowers zᴿ." `
        'Ineligible / NZ type N' "Outside: bᴺ + ℓ, with bᴺ = 0 in code`r`nWork: w − hᴺ`r`nFlow surplus: zᴺ = w − hᴺ − bᴺ − ℓ`r`nNormalisation sets hᴺ = 0 post-shock." `
        'teal' 'aqua' 'The code labels h as «health»; empirically it is a residual differential net work-surplus wedge that may absorb composition, policy interactions and measurement.' | Out-Null

    Add-TwoColumnSlide $pres 'A job is found only if an offer arrives and is accepted' $T 1 `
        'Offer arrival' "o(s,c) = 1 − exp(−cs)`r`n`r`nc is the contact probability from the matching block; s is chosen search effort. The exponential form keeps offers in [0,1]." `
        'Acceptance' "a(z) = Λ[(z − z̄)/σₐ]`r`n`r`nJob-finding probability:`r`nf = a(z) × o(s,c).`r`nThreshold z̄ and dispersion σₐ are fixed, not estimated." `
        'teal' 'orange' 'Observed job finding is a product of search, contacts and acceptance. Aggregate transition rates cannot separately identify all three.' | Out-Null

    Add-TwoColumnSlide $pres 'Search effort solves a dynamic marginal-cost condition' $T 1 `
        'First-order condition' "κ s^η = β · a · c · exp(−cs) · [W′ − U′]`r`n`r`nMarginal cost rises with curvature η; the gain depends on acceptance, contact and next-period value surplus." `
        'Interpretation' "Higher benefits reduce W−U and acceptance; both lower the return to search. Poor matching lowers c. Anticipated policy expiry changes W′−U′ before the cash payment ends." `
        'teal' 'aqua' 'In the selected calibration η = 100. That pins s close to one, so the modelʼs «labour-supply» response is mostly an acceptance response.' | Out-Null

    Add-TwoColumnSlide $pres 'Employment exits are a smooth hazard of static flow surplus' $T 1 `
        'Implemented equation' "δⱼ = clamp[δ₀ + ξ + A Λ((z̄ₛ − zⱼ)/σₛ)]`r`n`r`nδ₀: baseline`r`nξ: common post shock`r`nA: endogenous component amplitude" `
        'Economic meaning' "A lower zⱼ moves a group into the steep region of the hazard. Benefits and hᴿ therefore complement each other: the effect of one depends on the level of the other." `
        'teal' 'orange' 'There is no explicit match-quality state or joint continuation rule. Near term: call this a smooth reduced-form employment-exit hazard.' | Out-Null

    Add-TwoColumnSlide $pres 'The matching block turns μ into contact opportunities' $T 1 `
        'Current firm side' "Firm surplus = y − w`r`nJ = (y−w)/[1−β(1−δfirm)]`r`nθ = [β μ J / vacancy cost]^(1/α)`r`nc = μ θ^(1−α)" `
        'Internal inconsistency' "Firm continuation uses δfirm = 1% weekly, while worker-side exits are about 4.8%–9.3%. Policy-induced exits do not reduce J or feed back into vacancy creation." `
        'teal' 'red' 'Choose: integrate the same exit hazard into firm value, or treat c/μ as an exogenous opportunity process and drop free-entry language.' | Out-Null

    Add-TwoColumnSlide $pres 'Bellman equations make the shock dynamic' $T 1 `
        'Employed value' "Wⱼ,t = (w − hⱼ,t) + β[(1−δⱼ,t)Wⱼ,t+1 + δⱼ,t Uⱼ,t+1]" `
        'Non-employed value' "Uⱼ,t = (bⱼ,t + ℓ) − C(sⱼ,t) + β[fⱼ,t Wⱼ,t+1 + (1−fⱼ,t)Uⱼ,t+1]" `
        'teal' 'aqua' 'Backward induction solves values and search over the announced shock horizon; forward simulation averages weeks 12–36.' | Out-Null

    Add-BulletSlide $pres 'Timing assumptions are economically important' $T 1 @(
        'Weekly model: T = 60; shock weeks 12–36 inclusive (25 weeks).',
        'Shock is unanticipated at onset, then its end date is known with certainty.',
        'The manuscriptʼs empirical window, announcement date and supplement payment window do not currently line up with the model averaging window.',
        'β = 0.99 weekly implies an annual discount factor near 0.59—far below conventional calibrations.',
        'The real policy was announced, paid, extended and tapered in stages; expectations should follow that history.'
    ) 'Create a date-to-model-week table and set βweek = βannual^(1/52). Re-run every quantitative result.' 'orange' | Out-Null

    Add-SectionSlide $pres '03' 'Calibration and what the model says' 'A close fit can coexist with severe mechanism uncertainty.' | Out-Null

    Add-CalibrationFlowSlide $pres $T | Out-Null

    Add-TableSlide $pres 'Core fixed and calibrated quantities in the benchmark' $T 2 `
        @('Object','Benchmark','How set','Interpretation / concern') `
        @(
            @('Weekly β','0.99','Fixed','Implies implausibly low annual discounting'),
            @('Wage / output','0.55 / 1.00','Fixed normalisation','No wage bargaining or wage distribution'),
            @('Benefit RR','0.28 → 0.56','Stylised doubling','Actual policy is a flat $550/fortnight plus means tests'),
            @('Search curvature η','100','Selected from restricted grid','Pins search close to one'),
            @('Exit threshold / dispersion','0.10 / 0.04','Selected from grid','Many near-ties have radically different predictions'),
            @('μ pre / post','0.0739 / 0.0677','Fit JFR moments','Post μ fitted to NZ JFR'),
            @('Differential wedge hᴿ','0.0368 = 6.69% wage','Fit AU post exits','Residual wedge, not identified health disutility')
        ) @(1.25,0.85,1.15,2.1) 'A manuscript calibration table should label every item fixed, externally calibrated, fitted or held out.' `
        'Source: 15_structural_separation_summary.csv and model code.' | Out-Null

    Add-ImageSlide $pres 'Benchmark fit: the held-out error is visible but modest' $T 2 (Join-Path $AssetDir 'model_fit.png') `
        'The fit bars are not standard errors; targets are treated as exact in the current code.' `
        'Source: model-15 summary CSV.' 55 110 850 355 | Out-Null

    Add-ImageSlide $pres 'Scenario results: benefit effects and pandemic wedges interact' $T 2 (Join-Path $AssetDir 'model_scenarios.png') `
        'Do not divide standalone bars by the full bar and call the result a «share».' `
        'Source: model-15 decomposition CSV.' 55 108 850 360 | Out-Null

    Add-TwoColumnSlide $pres 'The model reverses the manuscriptʼs current job-finding headline' $T 2 `
        'What the manuscript says' "Most of the labour response came from pandemic work disutility rather than benefit generosity; benefit effects were smaller in normal conditions." `
        'What model 15 says' "Benefit-only job-finding DiD: −1.50 pp.`r`nFull model DiD: −1.31 pp.`r`nCommon COVID only: +0.24 pp.`r`nDifferential wedge only: −0.05 pp." `
        'red' 'green' 'For job finding, the benefit channel is at least as large in the no-COVID counterfactual as in the full benchmark. Rewrite abstract, introduction and conclusion.' | Out-Null

    Add-TwoColumnSlide $pres 'For exits, complementarity—not «health versus benefits»—is the result' $T 2 `
        'Standalone DiD effects' "Benefit only: +1.688 pp`r`nCommon COVID only: ≈0.000 pp`r`nDifferential wedge only: +0.061 pp`r`nSum: +1.749 pp" `
        'Full environment' "Full exit DiD: +3.720 pp`r`nResidual interaction: 3.720 − 1.749 = +1.971 pp`r`nThe interaction exceeds every standalone wedge other than benefits." `
        'teal' 'orange' 'Compute all 2³ counterfactual combinations and report a Shapley or explicit interaction decomposition.' | Out-Null

    Add-ImageSlide $pres 'The benchmarkʼs job-finding mechanism is mainly acceptance' $T 2 (Join-Path $AssetDir 'model_margins.png') `
        'This mechanism is imposed partly by η = 100 and fixed acceptance parameters; it is not directly validated by applications or offer data.' `
        'Source: model-15 scenario CSV.' 55 110 850 355 | Out-Null

    Add-ImageSlide $pres 'Weak identification appears in the near-fitting parameter set' $T 2 (Join-Path $AssetDir 'model_identification.png') `
        'A single «best» grid point is selected by tiny numerical differences in a loss many specifications match almost exactly.' `
        'Source: model-15 restricted grid; plotted points have fitted loss < 0.01.' 170 103 620 382 | Out-Null

    Add-ImageSlide $pres 'The replacement-rate gradient file is a diagnostic—not a paper result' $T 2 (Join-Path $AssetDir 'replacement_rate_placeholder.png') `
        'Rebuild using statutory predicted entitlements and pre-displacement wages before using it as external validation.' `
        'Source: model-15 replacement-rate validation CSV; code marks rates as placeholders.' 160 108 640 360 | Out-Null

    Add-SectionSlide $pres '04' 'What is strong, what is fragile?' 'Separate useful structure from claims the current model cannot carry.' | Out-Null

    Add-ThreeCardsSlide $pres 'What should be retained' $T 3 `
        @('Common surplus index','Held-out prediction','Transparent outputs') `
        @(
            'Benefits and work costs enter the same economic object, creating interpretable nonlinear responses across margins.',
            'The restricted calibration reserves Australian post job finding rather than mechanically fitting every target.',
            'Self-contained Julia code writes fit, scenario, grid and validation CSVs that can feed a reproducible paper pipeline.'
        ) 'These are solid foundations for an interpretation model once empirical targets, time units and uncertainty are corrected.' | Out-Null

    Add-TwoColumnSlide $pres 'What the current aggregate moments cannot identify' $T 3 `
        'Search-and-finding block' "κ level and η curvature`r`nMatching/contact efficiency μ`r`nOffer-arrival responsiveness`r`nAcceptance threshold and dispersion`r`nDynamic value surplus" `
        'Exit block' "Background firm destruction`r`nVoluntary versus involuntary exits`r`nMatch-specific quality distribution`r`nWorker and firm continuation values`r`nDifferential work-surplus wedge" `
        'red' 'red' 'Add external moments or report a set of observationally admissible mechanisms. Six aggregate cells do not identify this many margins.' | Out-Null

    Add-TwoColumnSlide $pres 'The residual work wedge needs a neutral label' $T 3 `
        'Current normalisation' "Set hᴺ = 0.`r`nFit common NZ exit increase with ξ.`r`nFit hᴿ to the remaining Australian exit increase.`r`nAny R-specific omitted force loads into hᴿ." `
        'What it may include' "Health risk and caregiving`r`nIndustry/employer composition`r`nJobKeeper interactions`r`nBenefit administration`r`nPayroll reporting and sample selection" `
        'teal' 'orange' 'Call hᴿ−hᴺ a differential net work-surplus wedge unless independent moments identify health or work disutility.' | Out-Null

    Add-TwoColumnSlide $pres 'The current model is not a welfare model' $T 3 `
        'Missing insurance side' "No consumption or saving`r`nNo risk aversion or liquidity constraints`r`nNo household heterogeneity`r`nNo taxes, budget or marginal cost of funds" `
        'Missing externalities' "No epidemiological/contact externality`r`nNo congestion or job-rationing externality`r`nNo firm survival or vacancy externality`r`nNo value of improved match quality" `
        'red' 'red' 'Valid conclusion: the paper estimates behavioural transitions and interprets channels. Invalid conclusion: the supplement had positive or negative net social welfare.' | Out-Null

    Add-BulletSlide $pres 'Uncertainty must pass through both empirical and structural stages' $T 3 @(
        'Resample the person-week design using blocks and, where relevant, matched pairs.',
        'Recompute the empirical target vector and covariance matrix in every draw.',
        'Recalibrate the model; record failures, bound hits and optimiser status.',
        'For each draw, retain the full economically admissible parameter set—not only one tie-broken optimum.',
        'Report intervals combining sampling uncertainty with set/functional-form uncertainty.',
        'Show model fit and counterfactual ranges, not only point bars.'
    ) 'Microdata improve measurement, but the policy still supplies one aggregate shock. Structural precision cannot exceed empirical identification.' 'orange' | Out-Null

    Add-SectionSlide $pres '05' 'A submission-ready model roadmap' 'Make a credible minimal model first; treat richer welfare and equilibrium extensions as optional.' | Out-Null

    Add-TableSlide $pres 'Prioritised structural work before submission' $T 4 `
        @('Priority','Action','Why it is binding','Deliverable') `
        @(
            @('1','Recalibrate to redesigned person-week risk-set moments','Current exit targets inherit post-treatment selection','Frozen target table + covariance matrix'),
            @('2','Correct β, dates and actual benefit entitlements','Current dynamics and replacement rates are not policy-consistent','Date map + entitlement distribution'),
            @('3','Report an admissible parameter set','Numerical tie-breaking hides weak identification','Ranges for fit, held-out prediction and counterfactuals'),
            @('4','Complete 2³ / Shapley decomposition','Current scenarios are non-additive','Mechanism table with interactions'),
            @('5','Fix firm-exit consistency or simplify matching block','Free entry ignores the modelʼs own exits','Integrated J or exogenous contact interpretation'),
            @('6','Add numerical tests and reproducible environment','Grid silently omits failures; no project lockfile','Project/Manifest, driver, tests and logs')
        ) @(0.65,1.5,1.65,1.45) 'Stop/go rule: if the redesigned empirical cells are unstable, pause structural refinement and narrow the paperʼs causal claims.' | Out-Null

    Add-TwoColumnSlide $pres 'Minimal model for the main paper' $T 4 `
        'Keep in main text' "Two types and states`r`nFlow-surplus index`r`nSearch FOC and acceptance`r`nReduced-form smooth exit hazard`r`nMoment-to-parameter table`r`nFit + held-out prediction range`r`nFactorial/Shapley decomposition" `
        'Move to appendix' "Bellman derivations`r`nFull parameter grid`r`nOptimiser diagnostics`r`nAll shock-path sensitivities`r`nAlternative normalisations`r`nBootstrap/set-construction details`r`nAll subgroup validations" `
        'green' 'grey' 'Aim for a transparent quantitative interpretation section, not a second paper inside the empirical paper.' | Out-Null

    Add-TwoColumnSlide $pres 'Ambitious extension: explicit match quality and consistent vacancy creation' $T 4 `
        'State extension' "Each match draws quality x.`r`nWorker and firm continuation depend on x.`r`nJoint surplus Sⱼ(x) determines an endogenous continuation cutoff x*ⱼ.`r`nExit probability is F(x*ⱼ)." `
        'Equilibrium extension' "Firm value integrates over type and quality.`r`nVacancy free entry uses the same exit process.`r`nBenefits affect vacancy value through duration and destruction.`r`nValidate tightness against vacancy data." `
        'teal' 'aqua' 'This would justify «structural match-quality threshold» language. It is valuable, but not necessary if the current hazard is labelled modestly.' | Out-Null

    Add-TwoColumnSlide $pres 'Ambitious extension: insurance, households and public-health externalities' $T 4 `
        'Household block' "CRRA utility, assets and borrowing limits`r`nFlat benefits + taper/means tests`r`nHeterogeneous wages and household composition`r`nConsumption response and insurance value" `
        'Social block' "Benefit financing and marginal cost of funds`r`nContact intensity by employment state`r`nEpidemiological transmission externality`r`nVacancy/job-rationing effects`r`nState-contingent optimal benefit paths" `
        'orange' 'teal' 'This is the route to welfare. It materially expands scope and data requirements; do not let it delay a strong behavioural paper.' | Out-Null

    Add-BulletSlide $pres 'Additional moments with the highest identification value' $T 4 @(
        'Applications/search contacts or reported search hours—discipline κ and η.',
        'Offer receipt, rejection and reservation wages—discipline acceptance threshold and dispersion.',
        'Vacancies/applications by local market—discipline contact efficiency and tightness.',
        'Employment-exit reason, stand-down, firm closure and JobKeeper exposure—separate worker continuation from firm destruction.',
        'Duration dependence in job finding—tests value dynamics and search technology.',
        'Actual predicted replacement-rate changes—tests the benefit mechanism across policy intensity.',
        'Post-taper and expiry paths—validate expectations and forward-looking behaviour.'
    ) 'Pre-register which moments are fitted and which are held out. Validation is strongest when the held-out set is chosen before specification search.' 'teal' | Out-Null

    Add-TableSlide $pres 'How the structural section should be written in the paper' $T 4 `
        @('Subsection','Purpose','One main exhibit') `
        @(
            @('Purpose and scope','Interpret transition gaps; explicitly exclude welfare and full GE','One-paragraph mechanism map'),
            @('Environment','Define types, states, outside options, search, acceptance and exits','State-transition diagram'),
            @('Equations','Show Bellman equations, search FOC, finding probability and exit hazard','Four-equation box'),
            @('Identification','Map every fixed/fitted/held-out object to a source or moment','Parameter–moment table'),
            @('Fit and validation','Show data, benchmark and near-fitting range','Fit + held-out interval figure'),
            @('Counterfactuals','Report factorial/Shapley effects and underlying margins','Mechanism decomposition'),
            @('Limits','Partial equilibrium, residual wedge, no welfare, functional-form sensitivity','Short boxed limitations')
        ) @(1.15,2.7,1.55) 'The benchmark is a visual anchor; conclusions must be based on the admissible range.' | Out-Null

    Add-TwoColumnSlide $pres 'Claim discipline: what the model can and cannot support' $T 4 `
        'Defensible' "«Model 15 shows that benefits can reduce job finding through surplus and acceptance.»`r`n«Benefits and work-surplus wedges interact in employment exits.»`r`n«A range of mechanisms is consistent with the aggregate moments.»" `
        'Not yet defensible' "«Most of the response is health rather than benefits.»`r`n«The differential wedge is a health preference.»`r`n«The exit hazard is a structural match-quality cutoff.»`r`n«The supplement increased net welfare.»" `
        'green' 'red' 'Good modelling language distinguishes mechanism possibility, quantitative consistency and identified causal decomposition.' | Out-Null

    Add-BulletSlide $pres 'Discussion questions for the research team' $T 4 @(
        'Is the modelʼs central contribution job-finding mechanisms, employment-exit complementarity, or state-contingent policy? Choose one.',
        'Can the secure data identify applications, offer acceptance, exit reason, JobKeeper exposure or actual entitlement replacement rates?',
        'Should vacancy creation remain endogenous if its firm value does not share the exit process?',
        'What economically admissible restrictions rule out 40% wage-equivalent differential wedges?',
        'Which moments will be genuinely held out before the next calibration?',
        'Would a simpler reduced-form hazard model tell the same durable story more credibly?',
        'What quantitative conclusion survives across every near-fitting specification?'
    ) 'Recommended seminar stance: lead with what the model teaches; volunteer the identification limits before the audience does.' 'teal' | Out-Null

    Save-Presentation $pres $pptx $pdf @(1,6,10,20,24,27,33,39)
    $count = $pres.Slides.Count
    $pres.Close()
    return [PSCustomObject]@{ Deck='Structural'; Slides=$count; PowerPoint=$pptx; PDF=$pdf }
}

function Build-EmpiricalDeck {
    param($Ppt)
    $pptx = Join-Path $OutDir 'COVID_labour_supply_empirical_design_teaching_deck.pptx'
    $pdf = Join-Path $OutDir 'COVID_labour_supply_empirical_design_teaching_deck.pdf'
    $T = @('Question', 'Current design', 'Threats', 'Re-estimation', 'Paper')
    $pres = Initialise-Presentation $Ppt $pptx 'Empirical design for emergency unemployment assistance' 'Australians and New Zealanders in shared labour markets: promise, threats and a journal-ready redesign'

    Add-ThreeCardsSlide $pres 'What this deck is designed to teach' $T 0 `
        @('Understand the estimand','Audit the design','Build the next version') `
        @(
            'Separate eligibility for the March policy bundle, realised receipt and the $550 supplement as distinct causal objects.',
            'Trace selection, risk-set, timing, policy-interaction, measurement and one-shock inference threats.',
            'Specify a pre-policy cohort, person-week hazard design, conservative uncertainty package and decisive exhibit set.'
        ) 'The institutional comparison is genuinely promising. The empirical rebuild—not extra prose—is the gate to a strong public-economics submission.' | Out-Null

    Add-ThreeCardsSlide $pres 'Bottom line on the current empirical design' $T 0 `
        @('Why it is attractive','Why it is not ready','What would change the verdict') `
        @(
            'Linked weekly payroll, tax, visa and benefit records; large benefit change; ineligible New Zealand citizens living in the same Australian labour markets.',
            'Treatment uses realised future benefit receipt; samples condition on future non-employment; inference treats 48 cells as IID; JobKeeper and the policy bundle remain confounded.',
            'Stable estimates after a pre-policy cohort rebuild, correct risk sets, long pre-trends, design-specific inference and direct policy-interaction checks.'
        ) 'Current point estimates are useful targets and descriptive facts. They are not yet journal-ready causal estimates.' | Out-Null

    Add-TwoColumnSlide $pres 'Roadmap' $T 0 `
        'Part I — Learn the current design' "Question and estimand`r`nInstitutional comparison`r`nData and matching`r`nWeekly outcomes`r`nDiD and current estimates" `
        'Part II — Rebuild for submission' "Post-treatment selection`r`nRisk sets and person-week estimation`r`nOne-policy inference`r`nPolicy bundle and JobKeeper`r`nHeterogeneity, earnings and paper exhibits" `
        'teal' 'aqua' 'The deck treats shortcomings as design decisions with concrete repairs, not as a generic robustness checklist.' | Out-Null

    Add-SectionSlide $pres '01' 'Question, estimand and institutional experiment' 'A durable public-finance question comes before an estimator.' | Out-Null

    Add-TwoColumnSlide $pres 'The durable question is broader than one pandemic programme' $T 0 `
        'Policy question' "How does a large increase in unemployment assistance change movements into and out of employment when jobs are scarce, health risks are salient and search requirements are relaxed?" `
        'Why labour economists care' "Search and acceptance responses`r`nLiquidity versus moral hazard`r`nWorker versus firm-driven exits`r`nState dependence of UI effects`r`nInteractions with wage subsidies and conditionality" `
        'teal' 'aqua' 'Frame Australia–New Zealand as the identification design; frame emergency benefit design as the contribution.' | Out-Null

    Add-TableSlide $pres 'Three plausible estimands are currently mixed together' $T 0 `
        @('Estimand','Treatment definition','Interpretation','Status') `
        @(
            @('Eligibility / policy-bundle ITT','Nationality × pre-policy legal eligibility','Effect of exposure to the March JobSeeker expansion','Preferred main estimand'),
            @('Realised receipt effect','Received JobSeeker at any time in 2020','Effect among endogenous takers / job losers','Post-treatment; secondary only'),
            @('$550 supplement alone','Incremental cash amount holding other reforms fixed','Pure benefit-level effect','Not isolated by baseline design'),
            @('Instrumented receipt','Eligibility instrument for actual receipt','Local effect under exclusion/monotonicity','Possible only with credible first stage and exclusion')
        ) @(1.05,1.35,1.65,1.15) 'Recommended sentence: estimate exposure to Australiaʼs March 2020 JobSeeker expansion relative to ineligible New Zealand citizens in shared local labour markets.' | Out-Null

    Add-TwoColumnSlide $pres 'The comparison group is unusual—and economically informative' $T 0 `
        'Australian citizens' "Potentially eligible for JobSeeker and the supplement`r`nAlso affected by eligibility expansions, waivers and mutual-obligation changes`r`nActual receipt depends on job loss and take-up" `
        'New Zealand citizens on SCV' "Live and work in Australia`r`nGenerally ineligible for JobSeeker within the selected residency window`r`nFace many common local shocks but may differ in networks, tenure, mobility and selection" `
        'teal' 'aqua' 'Say «shared local shocks conditional on observables and parallel trends»—not «identical conditions».' | Out-Null

    Add-PolicyTimelineSlide $pres $T | Out-Null

    Add-TableSlide $pres 'The reduced-form treatment is a policy bundle' $T 0 `
        @('Component','Timing / exposure','Why it matters for outcomes','Can baseline isolate it?') `
        @(
            @('Coronavirus Supplement','Announced 22 Mar; paid from 27 Apr','$550/fortnight changes outside option','No—coincides with other changes'),
            @('Eligibility and means-test changes','Implemented in March','Changes who can receive JobSeeker','No—part of nationality-linked exposure'),
            @('Waiting-period / asset-test waivers','March 2020','Changes access speed and effective generosity','No'),
            @('Mutual obligations','Suspended then staged return','Changes search cost/requirements','Potential triple-difference variation'),
            @('JobKeeper wage subsidy','Announced 30 Mar','Directly affects retention and vacancies','Needs actual/predicted exposure'),
            @('Health restrictions','National + state timing','Affect demand, risk and feasible work','Common only imperfectly')
        ) @(1.15,1.15,1.55,1.35) 'Main causal label should be «JobSeeker expansion / March policy bundle» unless component-specific variation is demonstrated.' | Out-Null

    Add-SectionSlide $pres '02' 'What the current design does' 'Linked data, realised-receipt matching, selected weekly risk sets and a two-group DiD.' | Out-Null

    Add-TableSlide $pres 'Administrative data are a major comparative advantage' $T 1 `
        @('Source','Role in design','Key validation question') `
        @(
            @('Single Touch Payroll','Weekly pay/employment histories and transitions','Does no pay mean no employment? How are late/irregular reports handled?'),
            @('Visa / Home Affairs','Identify New Zealand citizens and residency tenure','Who is legally ineligible and continuously resident?'),
            @('DSS / DOMINO','Identify JobSeeker receipt','Receipt is post-treatment; can pre-policy statutory eligibility be reconstructed?'),
            @('Personal income tax','Income, occupation, partnership and annual outcomes','Can pre-period wages and actual replacement rates be built?'),
            @('BLADE / business tax','Industry and employer characteristics','Can JobKeeper eligibility/receipt and firm exit be linked?')
        ) @(1.35,1.85,2.25) 'The data can support a much stronger design if the raw-to-person-week construction layer is rebuilt and documented.' | Out-Null

    Add-TwoColumnSlide $pres 'Current sample construction uses future outcomes twice' $T 1 `
        'Current treatment/sample' "Australian treatment: received JobSeeker at any time in 2020.`r`nBoth groups: at least four consecutive weeks out of work in 2020.`r`nMatched sample ≈22,873 AU + 22,859 NZ." `
        'Current exit risk' "Employment-exit analysis is drawn from people known to experience non-employment later in 2020.`r`nA policy-induced exit can add someone to the sample; a prevented exit can remove them." `
        'teal' 'red' 'This changes both composition and the meaning of the denominators. It is not repaired by more controls or propensity-score matching.' | Out-Null

    Add-ImageSlide $pres 'Current weekly patterns are striking—but inherit selected denominators' $T 1 (Join-Path $AssetDir 'empirical_weekly_series.png') `
        'Use these as descriptive motivation while rebuilding the cohort and risk sets.' `
        'Source: cleared JFR/SR group CSVs in initial_RDD.' 55 110 850 350 | Out-Null

    Add-ImageSlide $pres 'Current pre/post cells imply large differential changes' $T 1 (Join-Path $AssetDir 'empirical_did_cells.png') `
        'Primary reporting should use percentage points, named baseline rates and uncertainty—not only «19%» or «64%».' `
        'Source: manuscript Tables 3–4.' 65 108 830 365 | Out-Null

    Add-TwoColumnSlide $pres 'The baseline two-group DiD is transparent' $T 1 `
        'Equation' "Ygt = α + β₁ Australian_g + β₂ Post_t`r`n      + β₃(Australian_g × Post_t) + εgt`r`n`r`nβ₃ = [AU post−pre] − [NZ post−pre]." `
        'Causal requirement' "Absent the policy exposure, the adjusted Australian weekly transition path would have changed like the comparable NZ path.`r`nNo other nationality-differential shock occurs at the same time." `
        'teal' 'orange' 'The equation is simple; the hard work is cohort definition, policy isolation and uncertainty for one treated contrast.' | Out-Null

    Add-TableSlide $pres 'Current headline estimates—and why precision is provisional' $T 1 `
        @('Outcome','Draft DiD','Baseline','Draft uncertainty','Immediate caution') `
        @(
            @('Job finding','−1.71 pp','AU pre 8.68% weekly','SE 0.82 pp; p=0.043','48 nationality-week cells; IID errors'),
            @('Payroll employment exits','+3.72 pp','AU pre 4.84% weekly','SE 1.40 pp; p=0.011','Selected exit sample; IID errors'),
            @('Earnings / match quality','Claimed no gain','Annual descriptive indices','No causal interval in reviewed script','Rebuild formal outcomes or remove claim'),
            @('JobKeeper-adjusted exits','+2.60 pp','Same selected series','Single month dummy','Not a measure of actual exposure')
        ) @(1.25,0.85,1.0,1.2,1.55) 'Do not lock the abstractʼs 19% and 64% claims until the redesigned denominators and inference are final.' | Out-Null

    Add-ThreeCardsSlide $pres 'Why the design is still worth rebuilding' $T 1 `
        @('Measurement','Counterfactual','Policy relevance') `
        @(
            'Weekly administrative transitions can measure flows and spell dynamics much more sharply than household surveys.',
            'NZ citizens share Australian local labour markets yet differ in benefit access—a rare comparison for an aggregate emergency reform.',
            'The flat, large supplement generates meaningful variation in predicted replacement-rate changes and household exposure.'
        ) 'If the revised results survive, the design can speak to state-contingent unemployment assistance beyond COVID-19.' | Out-Null

    Add-SectionSlide $pres '03' 'Threats to identification and inference' 'The central issues are selection, one-shock uncertainty and bundled interventions.' | Out-Null

    Add-ColliderSlide $pres $T | Out-Null

    Add-TwoColumnSlide $pres 'Realised receipt is an outcome of the same process being studied' $T 2 `
        'Why it is endogenous' "Receipt depends on job loss, take-up, claim timing, partner income, assets and administration.`r`nJob finding and employment exits directly affect the chance of becoming a recipient." `
        'Preferred treatment' "Citizenship × pre-policy legal eligibility.`r`nIf feasible, predict 1 March eligibility from only pre-period information.`r`nUse actual receipt as a mediator or instrumented secondary analysis." `
        'red' 'green' 'Eligibility-based ITT sacrifices a «treated recipients» label but gains a coherent causal estimand.' | Out-Null

    Add-TwoColumnSlide $pres 'One policy shock does not create 48 independent experiments' $T 2 `
        'Current problem' "Weekly outcomes are serially correlated.`r`nTreatment turns on once for one nationality group.`r`nIID standard errors treat cell residuals as independent.`r`nClustering on people cannot create more policy shocks." `
        'Implication' "The effective design uncertainty is about the Australian-minus-NZ time series around one event.`r`nConfidence intervals may be materially wider than conventional cell-level or person-clustered errors." `
        'red' 'orange' 'Administrative N improves outcome measurement and adjustment; it does not solve few-policy inference.' `
        'Bertrand, Duflo & Mullainathan (2004); Conley & Taber (2011).' | Out-Null

    Add-TableSlide $pres 'A transparent inference package for one treated contrast' $T 2 `
        @('Method','What it preserves','Role / caveat') `
        @(
            @('Pre/post collapse','Low-frequency policy contrast','Conservative benchmark; few observations'),
            @('HAC on weekly AU−NZ gap','Serial correlation over calendar time','Show bandwidth and small-sample sensitivity'),
            @('Block bootstrap','Within-series dependence','Resample calendar/residual blocks, not individual weeks'),
            @('Pair/design bootstrap','Matching and person composition','Rebuild weights/risk sets within each draw'),
            @('Placebo-date distribution','How unusual the break is in admissible pre-periods','Reference distribution; dates are not literally random'),
            @('Multiple placebo years','Seasonality and recurring nationality gaps','Must use identical cohort/risk-set algorithm')
        ) @(1.35,1.65,2.1) 'Report intervals from several defensible methods. Do not use significance stars as the core argument.' | Out-Null

    Add-BulletSlide $pres 'A short flat pre-period has low power against damaging trend differences' $T 2 @(
        'Extend consistently into 2019 as far as payroll coverage permits.',
        'Plot raw and composition-adjusted Australian-minus-NZ gaps.',
        'Estimate event-study coefficients and joint pre-period tests, but do not treat non-rejection as proof.',
        'Report the minimum detectable pre-trend and how much differential trend overturns the conclusion.',
        'Apply Rambachan–Roth / HonestDiD sensitivity to bounded deviations from parallel trends.',
        'Use nationality-specific trends only as a sensitivity check; short-window extrapolation is model-dependent.'
    ) 'The identifying assumption is a counterfactual path claim. Balance and a pre-trend p-value are evidence—not validation.' 'orange' `
        'Roth (2022); Rambachan & Roth (2023).' | Out-Null

    Add-TwoColumnSlide $pres 'JobKeeper is a direct retention treatment, not an April nuisance dummy' $T 2 `
        'Why it confounds' "JobKeeper preserves worker-firm links, including zero-hours jobs.`r`nIt changes measured exits, vacancy creation and competition for new jobs.`r`nIts interaction with JobSeeker may differ by nationality." `
        'Preferred strategies' "Link actual receipt / firm eligibility.`r`nSplit supported and unsupported jobs.`r`nUse pre-pandemic predicted exposure or firm eligibility designs.`r`nRestrict to clearly ineligible firms as a bounded sample." `
        'red' 'green' 'If exposure cannot be measured, label residual confounding honestly—especially for employment exits.' | Out-Null

    Add-TwoColumnSlide $pres 'Payroll records are powerful, but missing pay is not always non-employment' $T 2 `
        'Potential false transitions' "Fortnightly or irregular payment cadence`r`nLate reports and corrections`r`nZero-hours retention`r`nEmployer STP entry/exit`r`nConcurrent jobs`r`nSelf-employment not recorded" `
        'Validation design' "1-, 2- and 4-week persistence`r`nFortnightly aggregation`r`nRegular-payer restriction`r`nQuarterly/annual tax cross-checks`r`nAny-job versus main-job states`r`nEmployer and person exit flags" `
        'orange' 'green' 'Call the outcome «exit from recorded payroll employment» unless voluntary quits and layoffs can be separated.' | Out-Null

    Add-TwoColumnSlide $pres 'Nationality comparability remains the substantive counterfactual assumption' $T 2 `
        'Plausible shared shocks' "Same national macroeconomy`r`nSame state/local restrictions`r`nOften same industries and occupations`r`nCommon payroll measurement system" `
        'Possible differential responses' "Migration tenure and networks`r`nHousehold composition`r`nReturn-migration options`r`nIndustry/employer sorting`r`nRisk tolerance and work orientation`r`nAccess to other support" `
        'teal' 'orange' 'Match/reweight on observables and pre-outcomes; test heterogeneity and mobility; state conditional parallel trends explicitly.' | Out-Null

    Add-TableSlide $pres 'Matching needs diagnostics, not only «balanced distributions»' $T 2 `
        @('Required diagnostic','What it reveals','Decision rule') `
        @(
            @('Standardised mean differences','Balance for every covariate and key interaction','Pre-specify acceptable threshold'),
            @('Variance ratios','Distributional imbalance beyond means','Report before/after'),
            @('Propensity overlap + trimmed share','Where counterfactual support exists','Restrict to common support'),
            @('Pre-policy outcome balance','Whether transition histories align','Include employment, JFR, exits, duration, earnings'),
            @('Caliper / replacement sensitivity','Dependence on matching algorithm','Compare with overlap/entropy weights'),
            @('Matched-pair identifiers','Dependence and design structure','Retain in estimation and bootstrap')
        ) @(1.55,2.0,1.55) 'Do not select the matching method by the most significant treatment effect. Matching improves comparability; it does not randomise nationality.' | Out-Null

    Add-SectionSlide $pres '04' 'A journal-ready re-estimation plan' 'Pre-policy cohorts, contemporaneous hazards, person-week models and pre-specified validation.' | Out-Null

    Add-RiskSetSlide $pres $T | Out-Null

    Add-TwoColumnSlide $pres 'Preferred person-week specifications' $T 3 `
        'Linear probability baseline' "Yᵢ,t+1 = week FE + β(AUᵢ×Postₜ)`r`n+ pre-X weights/interactions + duration FE + εᵢt`r`n`r`nEasy to interpret in percentage points; estimate on weekly risk sets." `
        'Discrete-time hazard complement' "Logit or complementary log-log link`r`nFlexible duration dependence`r`nCalendar-week effects`r`nMatched-pair or weighting structure`r`nAverage marginal effects for comparability" `
        'teal' 'aqua' 'Cell-level AU−NZ series remains essential for transparency and time-series inference; person-week estimation adds covariate and duration control.' | Out-Null

    Add-BulletSlide $pres 'Covariates should respect timing and risk-set mechanics' $T 3 @(
        'Calendar-week fixed effects.',
        'Only pre-treatment demographic, household, industry, occupation, earnings and attachment measures.',
        'Non-employment-duration fixed effects for job finding.',
        'Job tenure or pre-policy attachment controls for employment exit where available.',
        'Matched-pair fixed effects or design weights if matching remains central.',
        'Region-by-week and industry-by-week controls only after checking they do not absorb the sole useful contrast.',
        'Payroll cadence and employer coverage indicators.',
        'Exact person, employer, person-week and event counts in every table.'
    ) 'The preferred specification should be chosen for design coherence—not because it maximises β/SE.' 'teal' | Out-Null

    Add-TwoColumnSlide $pres 'Event-study design should diagnose dynamics, not certify parallel trends' $T 3 `
        'Main figure' "Weekly AU−NZ adjusted gap`r`nOne omitted pre-period`r`nLong 2019 pre-window`r`nAnnouncement and payment markers`r`nConfidence bands using design-specific uncertainty" `
        'Sensitivity overlays' "Alternative cohort/risk-set definitions`r`nHonestDiD bounds`r`nJobKeeper exposure split`r`nLeave-one-industry/region/age-out`r`nFortnightly aggregation`r`nPost-taper and expiry dynamics" `
        'teal' 'aqua' 'A good event study makes the identifying comparison visually auditable week by week.' | Out-Null

    Add-TwoColumnSlide $pres 'Mutual obligations require a direct pooled triple difference' $T 3 `
        'Current approach' "Run separate Victoria and non-Victoria DiDs around July.`r`nOne insignificant coefficient and one significant coefficient do not imply the coefficients differ." `
        'Required test' "Estimate AU × Post × NonVictoria with all lower-order terms.`r`nTest the triple interaction directly.`r`nPlot dynamics and confront Victoriaʼs distinct outbreak/restrictions." `
        'red' 'green' 'Interpret modestly: differential nationality gaps when requirements resumed—not a clean proof of «motivation» or zero conditionality effects.' | Out-Null

    Add-TwoColumnSlide $pres 'Heterogeneity should use actual predicted replacement-rate changes' $T 3 `
        'Current problem' "Income-bin DiDs are −2.37, −4.25 and −6.04 pp from low to high income.`r`nThe model file relabels these as low/mid/high replacement rates using placeholder rates.`r`nThat mapping is not valid." `
        'Required construction' "Predict statutory pre/post entitlement from pre-policy household information.`r`nDivide by a clearly defined pre-displacement wage.`r`nEstimate continuous/bin/spline interactions and direct equality/trend tests." `
        'red' 'green' 'Show absolute benefit changes and proportional replacement-rate changes; separate them from sector, age and household exposure.' | Out-Null

    Add-TwoColumnSlide $pres 'The earnings / match-quality claim needs a real treatment-effect design' $T 3 `
        'To retain the claim' "Unconditional earnings including zeros`r`nConditional re-employment wages with selection caveat`r`nQuarterly/weekly dynamics`r`n2021–22 outcomes`r`nEmployer/occupation quality`r`n13/26/52-week retention`r`nConfidence intervals and selection bounds" `
        'If not feasible' "Remove «no subsequent earnings gain» from abstract, introduction and conclusion.`r`nKeep descriptive indices only as secondary context with limited language." `
        'teal' 'orange' 'A null descriptive pattern is not evidence of no causal match-quality benefit.' | Out-Null

    Add-TableSlide $pres 'Robustness should target specific failure modes' $T 3 `
        @('Threat','High-value check','What would worry us') `
        @(
            @('Seasonality / break mining','Multiple pseudo-dates and 2019/2022 placebos','Similar breaks occur often'),
            @('Outcome definition','1/2/4-week persistence; fortnightly aggregation','Effect appears only under one cadence'),
            @('Policy timing','22 Mar, 24 Mar, early Apr, 27 Apr; anticipation windows','Estimate depends on undocumented cut-off'),
            @('Composition','Balanced/open cohorts; leave-one-group-out','One industry/region/age drives result'),
            @('JobKeeper','Actual/predicted exposure splits','Exit effect concentrated in supported jobs'),
            @('Matching','Calipers, replacement, overlap/entropy weights','Poor overlap or sign instability'),
            @('Parallel trends','Long pre-period + HonestDiD','Small plausible drift overturns result')
        ) @(1.3,2.4,1.5) 'Summarise a pre-specified battery in a robustness table/specification curve; avoid dozens of selectively narrated regressions.' | Out-Null

    Add-TableSlide $pres 'The main paper needs a small set of decisive empirical exhibits' $T 3 `
        @('#','Exhibit','Question answered') `
        @(
            @('1','Institutional timeline + eligibility diagram','Who is exposed to which policy component and when?'),
            @('2','Cohort and balance table + Love plot','Are groups comparable on pre-policy variables and outcomes?'),
            @('3','Long event-study gap figure','Do paths align before the event; when does the gap open and close?'),
            @('4','Main person-week effects table','What is the effect under exact risk sets and conservative inference?'),
            @('5','Threats / robustness table','How sensitive are results to JobKeeper, dates, trends and transition definitions?'),
            @('6','Predicted replacement-rate figure','Does response scale with actual policy intensity?'),
            @('7','Model fit / held-out range / interaction decomposition','Which mechanisms are quantitatively consistent with the revised moments?')
        ) @(0.35,2.4,2.45) 'Everything else—matching variants, placebo grids, subgroup tables and model derivations—can move to the online appendix.' | Out-Null

    Add-SectionSlide $pres '05' 'From analysis to a rewritten paper' 'Narrow the claim, lead with uncertainty and make every design choice auditable.' | Out-Null

    Add-TwoColumnSlide $pres 'A defensible contribution has three parts' $T 4 `
        'Empirical design and result' "Linked weekly administrative records compare pre-policy-similar Australian and ineligible New Zealand citizens in shared labour markets; estimate transitions with design-specific uncertainty and explicit bundle caveats." `
        'Mechanism and broader lesson' "A partial-equilibrium dynamic model shows how a flat benefit increase and residual work-surplus wedge can interact; emergency benefit effects need not equal normal-times elasticities." `
        'green' 'teal' 'Do not oversell «clean identification», «identical conditions», voluntary separations, isolated supplement effects or welfare.' | Out-Null

    Add-TableSlide $pres 'Recommended paper architecture' $T 4 `
        @('Section','Job') `
        @(
            @('1. Introduction','Question, comparison, revised estimates with intervals, model range, scope limits'),
            @('2. Institutional setting and timeline','Eligibility, bundle, JobKeeper, obligations, announcement/payment dates'),
            @('3. Data and cohort construction','Pre-policy universe, linkage, weekly states, risk sets, balance and coverage'),
            @('4. Empirical strategy and inference','Estimand, person-week hazards, time-series uncertainty, HonestDiD'),
            @('5. Effects on transitions','Job finding and recorded payroll exits; dynamics and magnitudes'),
            @('6. Mechanisms and robustness','Policy interactions, replacement rates, earnings, placebos'),
            @('7. Dynamic interpretation model','Model 15 rewritten, partial identification and interaction decomposition'),
            @('8. Conclusion','One broad lesson, uncertainty and explicit limits')
        ) @(1.55,4.45) 'Rename the «RDD» material as controlled interrupted time-series / event-time robustness; time is not locally random at the pandemic cut-off.' | Out-Null

    Add-TwoColumnSlide $pres 'Results-writing discipline' $T 4 `
        'Use' "Percentage points first`r`nBaseline rate and named denominator`r`nConfidence interval in prose`r`n«Estimated» / «consistent with»`r`nRecorded payroll-employment exit`r`nPolicy exposure / JobSeeker expansion`r`nFitted versus held-out model moments" `
        'Avoid' "Significance stars as evidence`r`nNon-significance = zero`r`nComparing subgroup significance`r`n«Identical conditions»`r`nVoluntary separation without exit reason`r`nSupplement-only causal label`r`nHealth or welfare labels not identified" `
        'green' 'red' 'A careful vocabulary will strengthen rather than weaken the paper: it tells referees the authors understand the designʼs exact reach.' | Out-Null

    Add-TableSlide $pres 'Execution sequence and submission gates' $T 4 `
        @('Gate','Work','Pass condition','If it fails') `
        @(
            @('G1 Cohort/risk sets','Rebuild without future receipt or non-employment','Effects remain economically meaningful','Reframe as descriptive or redesign'),
            @('G2 Inference','One-policy serial-correlation package','Intervals/placebos still support central claim','Do not present conventional 5% causality'),
            @('G3 Policy isolation','JobKeeper + policy bundle tests','Stable estimates under credible exposure controls','Label bundle; narrow exit claim'),
            @('G4 Model discipline','Recalibrate revised moments; admissible set','Qualitative mechanism survives near-fitting set','Use model only illustratively'),
            @('G5 Rewrite','New outline, consistent claims and numbers','Internal referee can state contribution in two sentences','Delay submission')
        ) @(1.05,1.7,1.7,1.45) 'Do not begin final prose polish until G1–G3 are passed; otherwise every headline number and structural target may move.' | Out-Null

    Add-BulletSlide $pres 'Likely seminar and referee questions' $T 4 @(
        'Why should Australian and New Zealand citizens have parallel counterfactual paths during a border closure and public-health shock?',
        'How can treatment be future JobSeeker receipt if receipt depends on the labour-market outcome?',
        'Why does the employment-exit sample condition on eventually becoming non-employed?',
        'What is the effective number of policy shocks, and why are the standard errors valid?',
        'How do JobKeeper, eligibility waivers and mutual obligations map into the treatment?',
        'Does a missing STP payment represent unemployment, zero hours or reporting cadence?',
        'Why do high-income bins show the largest JFR response under a flat supplement?',
        'Which model moments are truly held out and what survives across near-fitting specifications?'
    ) 'A strong paper answers these questions in the design and exhibit choices—before the referee report.' 'orange' | Out-Null

    Add-ThreeCardsSlide $pres 'Take-home messages' $T 4 `
        @('Keep','Rebuild','Reframe') `
        @(
            'The citizenship-based institutional comparison, linked weekly administrative data and focus on employment transitions.',
            'Treatment assignment, cohort/risk sets, person-week estimation, long pre-trends, one-shock inference, JobKeeper exposure and actual replacement rates.',
            'Emergency income support and employment transitions under weak demand—not a clean supplement-only welfare verdict.'
        ) 'If the redesigned results remain stable, this can become a compelling labour/public economics paper. The next marginal hour belongs in cohort and inference code, not extra robustness prose.' | Out-Null

    Save-Presentation $pres $pptx $pdf @(1,6,10,13,19,21,29,37,42)
    $count = $pres.Slides.Count
    $pres.Close()
    return [PSCustomObject]@{ Deck='Empirical'; Slides=$count; PowerPoint=$pptx; PDF=$pdf }
}

$ppt = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = -1
    $results = @()
    $results += Build-StructuralDeck $ppt
    $results += Build-EmpiricalDeck $ppt
    $results | Format-Table -AutoSize
}
finally {
    if ($ppt) { $ppt.Quit() }
}
