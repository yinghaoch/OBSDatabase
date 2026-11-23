# OBS 紀錄

## OBS ID
```text
SIO2415657
```

## Short Description
```text
[SIT][WHQL][OB3 ABEI][WAVE1][RPLU][25H2]Touch pad Tests is Failed
```

## Long Description
```text
Test Platform: ABEI
Stage: SI
Image/OS-patch; Language: Win11 Professional(64) (26WW0AIZ6ae/EN)
BIOS:B.04
Test tool: HLK(10.1.26100.4946
Test unit: SKU12,SKU13
Fail unit: 2 out of 2 Units
Touch pad: SYNA
Description:
Test.InputResolution
Test.ReportRate
Test.ReportRateMultiple
Test.StationaryJitter
Test.StationaryJitterMultiple

Please check detail log in attachment.
```

## Steps to Reproduce
```text
01. Dash image (26WW0AIZ6ae/EN) with legacy enable, then OOBE image with Legacy disable.
02. Install the "HLK CLIENT" to the unit being tested
03. Execute "HLK STUDIO" on HLK Controller
04. On the Project tab, click Create project to create new project.
05. Create a machine Pool in HLK Configuration
06. Select and drag the test unit from Default Pool to created machine Pool
07. Change machine status to Ready
08. On the Selection tab,choosed "systems" then selected test unit.
09. On the Test tab,selected " Touch pad Tests " then clicked "Run Selected"
10. Wait and the test jobs will be completed.
```

## Created
```text
2026-04-17 17:18:24
```

## 附件

無附件。
