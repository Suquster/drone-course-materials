# 智能无人飞行器 - 实验代码与模型文件

> **学生**: 莫仁鹰  
> **学号**: 2023212167  
> **课程**: 智能无人飞行器的设计与应用

---

## 目录结构

```
实验代码/
├── yolov5-training/          # YOLOv5训练工程 (实验6.1 + 6.2)
│   ├── train.py              # 模型训练脚本 (已加学号标识)
│   ├── detect.py             # 目标检测脚本 (已加学号水印功能)
│   ├── export.py             # 模型导出脚本 (已加学号标识)
│   ├── data/ccsszz.yaml      # 数据集配置 (已加学号标识)
│   ├── datasets/             # 数据集 (613训练+218验证，已标注)
│   ├── yolov5s.pt            # 预训练基础权重
│   └── runs/train/           # 训练结果
│       ├── exp20/weights/best.pt   # ★ 已训练好的最佳模型
│       ├── exp5/weights/best.pt
│       └── exp4/weights/best.pt
│
├── yolov5-export/            # YOLOv5导出工程 (实验6.3)
│   └── yolov5-master/yolov5-master/
│       ├── best.pt           # 已训练好的模型
│       ├── best.onnx         # ★ 已转换好的ONNX模型
│       ├── export.py         # PT→ONNX导出脚本 (已加学号标识)
│       └── detect.py         # 检测脚本 (已加学号标识)
│
├── convert_onnx_to_rknn.py   # ★ ONNX→RKNN一键转换脚本 (虚拟机中使用)
└── README.md                 # 本文件
```

---

## 已完成的工作

### 实验 6.1 — 图像数据标注与YOLO数据集构建
- [x] 613张坦克训练图片已标注 (`datasets/images/train/` + `datasets/labels/train/`)
- [x] 218张坦克验证图片已标注 (`datasets/images/val/` + `datasets/labels/val/`)
- [x] 数据集配置文件 `data/ccsszz.yaml` 已配置完成

### 实验 6.2 — 训练环境搭建与模型训练
- [x] YOLOv5训练环境已搭建
- [x] 模型已训练完成，最佳权重: `runs/train/exp20/weights/best.pt`
- [x] 训练命令:
  ```
  python train.py --data data/ccsszz.yaml --weights yolov5s.pt --epochs 300 --batch-size 32 --imgsz 640 --optimizer Adam --name exp_2023212167
  ```

### 实验 6.3 — 模型格式转换与嵌入式部署
- [x] **PT → ONNX** 转换已完成: `best.onnx` (28MB)
- [ ] **ONNX → RKNN** 转换需在Ubuntu虚拟机中执行 (见下方步骤)
- [ ] **部署到RK3566** 需连接实际开发板

### 学号标识 (加分项)
所有关键代码文件均已添加学号标识:
- `detect.py` — 检测结果图片**自动显示学号水印** "2023212167 MoRenYing"
- `train.py` — 文件头部注释含学号
- `export.py` — 文件头部注释含学号
- `data/ccsszz.yaml` — 数据集配置含学号
- `convert_onnx_to_rknn.py` — 转换脚本含学号

---

## 你需要做的操作

### 步骤1: ONNX → RKNN 转换 (在Ubuntu虚拟机中)

1. 打开VMware，启动课程提供的Ubuntu虚拟机
2. 将 `yolov5-export/yolov5-master/yolov5-master/best.onnx` 复制到虚拟机中:
   ```
   rknn-toolkit2-master/examples/onnx/yolov5/
   ```
3. 确保该目录下有 `tank.jpg` 测试图片和 `dataset.txt` (内容为 `tank.jpg`)
4. 打开终端，执行:
   ```bash
   cd rknn-toolkit2-master/examples/onnx/yolov5/
   conda activate rknn
   python test.py
   ```
5. 等待转换完成，生成 `best.rknn`

> 也可以使用我准备的 `convert_onnx_to_rknn.py` 脚本（功能相同，多了学号标识输出）

### 步骤2: 部署到RK3566开发板

1. 通过MobaXterm连接RK3566开发板
2. 将 `best.rknn` 上传到开发板指定路径
3. 启动摄像头节点:
   ```bash
   roslaunch rknn_ros camera.launch device:=video0
   ```
4. 启动YOLO识别:
   ```bash
   roslaunch rknn_ros yolov5.launch chip_type:=RK356X
   ```
5. 查看识别效果:
   ```bash
   rqt_image_view
   # 选择 /rknn_image/theora 查看识别后的效果
   ```

### 步骤3: 验证检测结果 (可选，在训练环境中)

运行detect.py验证模型效果（检测结果图片会自动带学号水印）:
```bash
cd yolov5-training
python detect.py --weights runs/train/exp20/weights/best.pt --source datasets/images/val
```
结果保存在 `runs/detect/exp*/` 目录下，每张图片右下角显示 "2023212167 MoRenYing"。

---

## 关键文件说明

| 文件 | 说明 | 是否现成 |
|------|------|---------|
| `best.pt` | 训练好的坦克检测模型 (PyTorch格式) | 现成可用 |
| `best.onnx` | ONNX中间格式模型 | 现成可用 |
| `best.rknn` | RK3566 NPU加速模型 | 需在虚拟机中转换 |
| `yolov5s.pt` | YOLOv5s预训练基础权重 | 现成可用 |
| `ccsszz.yaml` | 坦克数据集配置 | 现成可用 |

---

*莫仁鹰 2023212167*
