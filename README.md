# Endfield_Simulation
终末地抽卡模拟

这是一个基于蒙特卡洛方法的抽卡策略分析项目。
通过 Python 模拟从开服到三周年的抽卡（共104个卡池），对比不同策略下的收益。

详细的策略分析请观看我的 Bilibili 视频：https://www.bilibili.com/video/BV1GvF7zpEEa

本项目代码使用了 AI 辅助，进行逻辑实现与图表绘制优化。

## 如何运行
修改开头的变量 SIMULATION_COUNT 来设定模拟次数，initial_permit_default 来设定初始抽数。
需要安装 Python 3 以及以下库：
```bash
pip install xlsxwriter


