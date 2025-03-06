## 信用卡账单分析工具

本工具可将微信、美团、京东、支付宝等支付平台导出的账单，自动关联到信用卡账单。

使用方式：
```
./billmerge -main 信用卡账单.xlsx 关联账单1.xlsx 关联账单2.xlsx ...
```

注意：
1、关联账单的第一行会被忽略，前三列应该为：日期、说明、金额，日期格式为 YYYY-MM-DD，如：2022-01-01
