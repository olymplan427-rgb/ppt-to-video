param(
  [Parameter(Mandatory=$true)][string]$PptxPath,
  [Parameter(Mandatory=$true)][string]$DurationsJson,
  [Parameter(Mandatory=$true)][string]$OutDir,
  [double]$EffectDuration = 0.5,
  [double]$Stagger = 0.25,
  [int]$Width = 1920,
  [int]$Fps = 30
)
# PowerPoint 자체 렌더러로 슬라이드에 요소별 순차 페이드인을 입혀 슬라이드마다 개별 mp4로
# 내보낸다. HTML로 재현하지 않고 원본 폰트·사진·레이아웃 그대로 유지하면서 모션만 더하는 것이
# 목적 — 1차 제작 시도가 실패한 지점(HTML 재현으로 브랜드 룩이 어긋남)을 다시 밟지 않기 위함.
#
# 슬라이드를 하나씩만 보이게(Hidden 토글) 하고 CreateVideo를 호출하면 숨긴 슬라이드는 출력에서
# 완전히 빠진다(실측 확인됨) — 슬라이드마다 파일을 복사·재오픈할 필요 없이 한 세션에서 처리한다.

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
New-Item -ItemType Directory -Force $OutDir | Out-Null
$durations = Get-Content $DurationsJson -Raw | ConvertFrom-Json

$app = New-Object -ComObject PowerPoint.Application
$pres = $app.Presentations.Open($PptxPath, $false, $false, $false)

foreach ($prop in $durations.PSObject.Properties) {
  $slideNo = [int]$prop.Name
  $slide = $pres.Slides.Item($slideNo)
  $seq = $slide.TimeLine.MainSequence
  $n = $slide.Shapes.Count

  # 값이 숫자 하나면 예전 방식(균등 예산), 객체({duration,checkpoints})면 나레이션
  # 문장이 실제로 발화되는 시점에 맞춘 구간별 등장 — 아래에서 분기.
  $val = $prop.Value
  $isObj = $val.PSObject.Properties.Name -contains 'duration'
  $duration = if ($isObj) { [double]$val.duration } else { [double]$val }
  $checkpoints = if ($isObj -and $val.checkpoints) { @($val.checkpoints | ForEach-Object { [double]$_ }) } else { $null }

  if ($checkpoints -and $checkpoints.Count -gt 0) {
    # 나레이션 문장 수만큼 구간을 나누고, 화면 요소를 문서 순서 그대로 그 구간에 배분한다.
    # 각 구간의 요소들은 그 구간이 끝나는 시각(=그 문장이 다 말해지는 시각)까지 순차
    # 페이드인을 마치도록 간격을 역산 — "전부 처음 2초에 등장, 나머지는 정적"이던
    # 문제를 없애고 화면이 나레이션을 따라 계속 살아있게 만든다.
    # 헤더(페이지 타이틀·슬라이드 메인 메시지·브랜드 로고)와 본문 요소를 나눈다.
    # 헤더는 씬이 열리면서 진행자 멘트와 함께 곧바로 뜬다 — 화제를 여는 동안 화면이
    # 비어 있으면 어색하고, 이 요소들은 전문가가 하나씩 짚어 설명할 대상이 아니다.
    # 본문 요소만 전문가 대사 체크포인트에 배분한다(사용자 지적 2026-08-23).
    # 판정은 위치로 한다 — 이 덱은 제목이 상단에, 로고가 상단 모서리에 고정돼 있다.
    $headerCut = $pres.PageSetup.SlideHeight * 0.30
    $headerIdx = @(); $bodyIdx = @()
    for ($i = 1; $i -le $n; $i++) {
      if ($slide.Shapes.Item($i).Top -lt $headerCut) { $headerIdx += $i } else { $bodyIdx += $i }
    }
    # 본문이 하나도 없으면(표지·간지 슬라이드) 전부 헤더로 취급해 한 번에 띄운다.
    if ($bodyIdx.Count -eq 0) { $headerIdx = 1..$n; $bodyIdx = @() }

    foreach ($i in $headerIdx) {
      $eff = $seq.AddEffect($slide.Shapes.Item($i), 10, 0, 3)
      $eff.Timing.Duration = 0.45
      $eff.Timing.TriggerDelayTime = 0.12     # 촘촘히 이어 붙여 도입부에 함께 등장
    }

    $nb = $bodyIdx.Count
    # 본문 요소가 체크포인트보다 적으면 예전엔 Min()으로 잘라 '앞쪽' 체크포인트만 썼다.
    # 그러면 설명이 한창일 때 화면은 이미 다 나와 정지한다. 잘라내지 말고 전 구간에서
    # 고르게 솎아, 마지막 요소가 마지막 문장에 떨어지게 한다.
    $cps = @($checkpoints)
    if ($nb -gt 0 -and $nb -lt $cps.Count) {
      $picked = @()
      for ($j = 1; $j -le $nb; $j++) {
        $picked += $cps[[int][Math]::Ceiling($j * $cps.Count / $nb) - 1]
      }
      $cps = $picked
    }
    $ngroups = [Math]::Min($cps.Count, [Math]::Max(1, $nb))
    $base = [Math]::Floor($nb / $ngroups)
    $rem = $nb % $ngroups
    $bi = 0
    $prevCp = 0.0
    for ($g = 0; $g -lt $ngroups; $g++) {
      if ($nb -eq 0) { break }
      $groupSize = $base + $(if ($g -ge ($ngroups - $rem)) { 1 } else { 0})
      if ($groupSize -lt 1) { continue }
      $target = $cps[$g]
      $window = [Math]::Max(0.15, $target - $prevCp)
      $gap = if ($groupSize -gt 1) { [Math]::Max(0.06, [Math]::Min(1.2, $window / ($groupSize - 1))) }
             else { [Math]::Min(0.7, $window) }
      $effDur = $gap * 0.6
      $effDelay = $gap * 0.4
      for ($k = 0; $k -lt $groupSize; $k++) {
        if ($bi -ge $nb) { break }
        $sh = $slide.Shapes.Item($bodyIdx[$bi])
        $eff = $seq.AddEffect($sh, 10, 0, 3)
        $eff.Timing.Duration = $effDur
        $eff.Timing.TriggerDelayTime = $effDelay
        $bi += 1
      }
      $prevCp = $target
    }
  } else {
    # 표 형태 슬라이드는 요소가 20개를 넘기도 해, 고정 스태거를 쓰면 전체 등장에 수십 초가
    # 걸린다(28개 x 0.75s ≈ 21초). 체크포인트가 없을 때의 대체 — 예산 안에서 균등 배분.
    $budget = [Math]::Min(2.2, 0.5 * $duration)
    $gap = if ($n -gt 1) { [Math]::Max(0.06, [Math]::Min(($EffectDuration + $Stagger), $budget / ($n - 1))) }
           else { $EffectDuration + $Stagger }
    $effDur = $gap * 0.6
    $effDelay = $gap * 0.4
    for ($i = 1; $i -le $n; $i++) {
      $sh = $slide.Shapes.Item($i)
      # msoAnimEffectFade=10, msoAnimTriggerAfterPrevious=3 — 앞 요소가 끝나야 다음이
      # 시작하는 진짜 순차 등장. (2=msoAnimTriggerWithPrevious는 전부 동시에 뜨는 것이라
      # 오작동이었음.) 튀는 모션(바운스·줌)은 클리셰 블랙리스트 취지에 안 맞아 단순
      # 페이드만 쓴다.
      $eff = $seq.AddEffect($sh, 10, 0, 3)
      $eff.Timing.Duration = $effDur
      $eff.Timing.TriggerDelayTime = $effDelay
    }
  }

  $slide.SlideShowTransition.AdvanceOnTime = 1
  $slide.SlideShowTransition.AdvanceTime = $duration
}

foreach ($i in 1..$pres.Slides.Count) { $pres.Slides.Item($i).SlideShowTransition.Hidden = 1 }

foreach ($prop in $durations.PSObject.Properties) {
  $slideNo = [int]$prop.Name
  $pres.Slides.Item($slideNo).SlideShowTransition.Hidden = 0
  $out = Join-Path $OutDir ("slide_{0:d2}.mp4" -f $slideNo)

  # CreateVideo가 산발적으로 실패한다(같은 슬라이드가 재실행하면 성공하는 경우도 관찰됨) —
  # 최대 2회 재시도한다. 기존에 잘 만들어진 $out은 새 렌더가 실제로 성공한 뒤에만
  # 덮어쓴다 — 먼저 지워버리면 재시도까지 실패했을 때 좋은 이전 버전마저 잃는다.
  # 시도마다 임시 파일명을 다르게 써서, 이전 시도의 파일이 아직 잠겨 있어도(안티바이러스
  # 스캔 등) 삭제하려다 스크립트 전체가 죽는 일이 없게 한다 — 삭제 실패는 청소 실패일
  # 뿐 치명적 오류가 아니므로 try/catch로 흡수한다.
  # 재시도 2회로는 부족했다 — 43장 렌더에서 5장(2·5·9·26·36)이 두 번 다 status=2로
  # 실패해 정지 이미지로 떨어졌다. 같은 슬라이드가 다음 실행에선 한 번에 성공하는 걸
  # 보면 콘텐츠 문제가 아니라 PowerPoint 인코더의 산발적 실패다. 4회로 늘린다.
  $status = 0
  $tmp = $null
  for ($attempt = 1; $attempt -le 4; $attempt++) {
    $tmp = "$out.attempt$attempt.tmp.mp4"
    try { if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction Stop } } catch { }
    $pres.CreateVideo($tmp, $true, 2, $Width, $Fps, 85)
    $deadline = (Get-Date).AddSeconds(120)
    do {
      Start-Sleep -Milliseconds 500
      $status = $pres.CreateVideoStatus
    } while ($status -eq 1 -and (Get-Date) -lt $deadline)
    if ($status -eq 3) { break }
    Write-Output "RETRY $slideNo attempt=$attempt status=$status"
  }

  if ($status -ne 3) {
    try { if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction Stop } } catch { }
    Write-Output "FAILED $slideNo status=$status"
  } else {
    # CreateVideoStatus=3은 "인코딩이 끝났다"는 뜻이지 "파일 핸들을 놓았다"는 뜻이 아니다.
    # PowerPoint가 잠깐 더 붙들고 있어서 곧바로 Move하면 IOException이 나고, 그게
    # 비종료 오류로 새어 나가 스크립트 종료코드를 1로 만들어 렌더 전체가 죽는다
    # (43장 렌더 중 slide_03에서 실제로 이 경합으로 중단됨). 놓을 때까지 기다린다.
    $moved = $false
    for ($m = 1; $m -le 30; $m++) {
      try { Move-Item -Force $tmp $out -ErrorAction Stop; $moved = $true; break }
      catch { Start-Sleep -Milliseconds 500 }
    }
    if ($moved) {
      Write-Output "RENDERED $slideNo $out"
    } else {
      Write-Output "FAILED $slideNo status=locked"
    }
  }
  $pres.Slides.Item($slideNo).SlideShowTransition.Hidden = 1
}

$pres.Close()
# 사용자가 다른 PPT를 열어 둔 채일 수 있다 — 우리가 연 것만 닫고, 아무것도 안 남았을 때만 종료.
if ($app.Presentations.Count -eq 0) { $app.Quit() }
