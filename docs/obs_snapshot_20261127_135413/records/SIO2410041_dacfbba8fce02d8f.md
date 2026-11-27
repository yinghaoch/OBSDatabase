# OBS 紀錄

## OBS ID
```text
SIO2410041
```

## Short Description
```text
[SIT][WHQL][25H2 SFR][KRK2]TPM 2.0 EK Certificate Tests failed
```

## Long Description
```text
Test Platform: HokieR25C1
Stage: PV
Image/OS-patch; Language: Win11 Professional(64)(25WW2HOZ6ac/EN)
BIOS: B.14
Test tool: HLK(10.1.26100.4946)
Test unit: SKU4
Fail Rate: 1 out of 1 Units


Description:
Run TPM 2.0 EK Certificate Tests failed

ErrorMessage:
Key ID: 02 2d db c1 11 39 8c 20 3b aa d8 b7 98 f0 f5 71 f6 70 3b 82 
An ECC NIST P384 EK certificate was not found. This type of EK certificate is required as per the Device.TrustedPlatformModule.TPM20.EKCert section of the WHCP spec as of October 2020. 
Please check detail log in attachment.
```

## Steps to Reproduce
```text
01. Dash/Unbundle image (CNB25WW2HOZ6ac /EN).
02. Install the "HLK CLIENT" to the unit being tested
03. Execute "HLK STUDIO" on HLK Controller
04. On the Project tab, click Create project to create new project
05. Create a machine Pool in HLK Configuration
06. Select and drag the test unit from Default Pool to created machine Pool
07. Change machine status to Ready
08. On the Selection tab,chose "systems" then selected test unit and load playlist
09. On the Test tab,selected "TPM 2.0 EK Certificate Tests" then clicked "Run Selected" 
10. Wait and the test jobs will be completed
```

## Created
```text
2026-04-10 17:41:17
```

## 附件

無附件。
