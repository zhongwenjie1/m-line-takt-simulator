# 本地打包说明

本阶段用于本地验证 PyInstaller onefile 打包，不提交 dist/build 产物。

## 环境

- 项目路径：/Users/zhongwenjie/程序/ticket_export_app
- Python：/Users/zhongwenjie/程序/.venv312/bin/python
- Python 版本：3.12.10
- PyInstaller：6.4.0
- 当前平台：macOS arm64
- 当前产物为 macOS arm64 可执行文件，不是 Windows exe

## 检查命令

```bash
bash ticket_export_app/scripts/build_local.sh check
```

## 打包命令

```bash
bash ticket_export_app/scripts/build_local.sh build
```

## 产物路径

```text
ticket_export_app/dist/生产指示小工具
```

## Spec 文件

```text
ticket_export_app/packaging/ticket_export_app.spec
```

## pandas / numpy 处理

由于程序依赖 pandas / numpy，spec 中显式收集：

- numpy submodules / data files / dynamic libs
- pandas submodules / data files / dynamic libs

首次打包产物运行时曾在 pandas -> numpy 导入阶段出现：

```text
RuntimeError: CPU dispatcher tracer already initlized
```

增加 numpy / pandas 收集配置后，产物可启动，未复现该错误。

## 已知 Warning

- pkg_resources deprecation warning
- numpy/pandas tests 子模块因 pytest 缺失未收集
- 少量可选 hidden import / Windows 库 warning

这些 warning 当前未阻止 macOS arm64 产物打包和启动。

## 不要提交

不要提交：

- ticket_export_app/dist/
- ticket_export_app/build/
- __pycache__
- .DS_Store

## Windows 打包

PyInstaller 不是跨平台交叉编译工具。Windows exe 需要在 Windows 环境打包。

后续可通过 GitHub Actions 的 windows-latest runner 打包 Windows exe。

## 后续计划

- 本地 macOS 打包继续验证
- 后续新增 GitHub Actions Windows 打包
- 后续考虑版本检查 + 下载提示
- 暂不做自动替换更新
