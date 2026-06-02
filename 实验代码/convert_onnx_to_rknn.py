# ============================================================
# ONNX → RKNN 模型转换脚本
# 学生: 莫仁鹰 (学号: 2023212167)
# 课程: 智能无人飞行器的设计与应用
# 用途: 将YOLOv5的ONNX模型转换为RK3566 NPU加速的RKNN格式
# 使用环境: Ubuntu虚拟机 + rknn-toolkit2 (conda activate rknn)
# ============================================================
#
# 使用方法:
#   1. 将此脚本和 best.onnx 放到虚拟机的 rknn-toolkit2-master/examples/onnx/yolov5/ 目录
#   2. 确保该目录下有一张坦克测试图片 (tank.jpg) 和 dataset.txt (内容为 tank.jpg)
#   3. 激活rknn环境: conda activate rknn
#   4. 运行: python convert_onnx_to_rknn.py
#   5. 生成的 best.rknn 即为部署文件
#
# 注意: 此脚本是对原有 test.py 的封装，添加了学号标识和更清晰的输出

import os
import sys
import numpy as np

def convert():
    try:
        from rknn.api import RKNN
    except ImportError:
        print("错误: 未安装 rknn-toolkit2")
        print("请先激活rknn环境: conda activate rknn")
        sys.exit(1)

    ONNX_MODEL = 'best.onnx'
    RKNN_MODEL = 'best.rknn'
    DATASET = './dataset.txt'

    if not os.path.exists(ONNX_MODEL):
        print(f"错误: 找不到 {ONNX_MODEL}")
        print("请将 best.onnx 文件复制到当前目录")
        sys.exit(1)

    print("=" * 60)
    print("ONNX → RKNN 模型转换")
    print("学生: 莫仁鹰 (学号: 2023212167)")
    print("=" * 60)

    # 创建RKNN对象
    rknn = RKNN()

    # 配置模型
    print("\n[1/4] 配置模型参数...")
    rknn.config(
        mean_values=[[0, 0, 0]],
        std_values=[[255, 255, 255]],
        target_platform='rk3566'
    )

    # 加载ONNX模型
    print("[2/4] 加载ONNX模型...")
    ret = rknn.load_onnx(model=ONNX_MODEL)
    if ret != 0:
        print("加载ONNX模型失败!")
        sys.exit(1)

    # 构建模型 (INT8量化)
    print("[3/4] 构建RKNN模型 (INT8量化)...")
    ret = rknn.build(do_quantization=True, dataset=DATASET)
    if ret != 0:
        print("构建RKNN模型失败!")
        sys.exit(1)

    # 导出RKNN模型
    print("[4/4] 导出RKNN模型...")
    ret = rknn.export_rknn(RKNN_MODEL)
    if ret != 0:
        print("导出RKNN模型失败!")
        sys.exit(1)

    print("\n" + "=" * 60)
    print(f"转换成功! 输出文件: {RKNN_MODEL}")
    print(f"文件大小: {os.path.getsize(RKNN_MODEL) / 1024 / 1024:.1f} MB")
    print("学号: 2023212167 | 姓名: 莫仁鹰")
    print("=" * 60)
    print("\n下一步: 将 best.rknn 部署到RK3566开发板")
    print("  scp best.rknn user@<RK3566_IP>:/path/to/deploy/")

    rknn.release()

if __name__ == '__main__':
    convert()
