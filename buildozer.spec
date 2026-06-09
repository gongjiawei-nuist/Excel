[app]

# (str) 你的软件在手机上显示的中文/英文名称
title = 发货计划计划管理

# (str) 软件的英文包名（只能是小写字母和数字，不能有特殊字符）
package.name = deliveryapp

# (str) 组织域名（用于生成安卓唯一的 Application ID）
package.domain = org.planmanager

# (str) 源代码所在目录，点（.）代表当前根目录
source.dir = .

# (list) 允许打包进手机的文件后缀。注意：这里我已经帮你把 xlsx 加进去了！
source.include_exts = py,png,jpg,kv,atlas,xlsx

# (str) 软件的版本号
version = 1.0

# (list) 核心依赖库！这里非常关键，我已经帮你配置好了 Python3、Kivy 以及处理 Excel 必须的 openpyxl
requirements = python3,kivy==2.3.0,openpyxl,et_xmlfile

# (str) 手机支持的显示方向 (landscape横屏, portrait竖屏, all自动旋转)
orientation = portrait

# (bool) 是否全屏显示 (0 为否，1 为是)
fullscreen = 1

# =============================================================================
# 安卓专属配置 (Android specifics)
# =============================================================================

# (list) 手机权限：这里开启了联网权限、以及最关键的【手机存储读写权限】，否则 Excel 无法保存到手机里
android.permissions = INTERNET, WRITE_EXTERNAL_STORAGE, READ_EXTERNAL_STORAGE, MANAGE_EXTERNAL_STORAGE

# (int) 目标安卓 API 版本（33 对应 Android 13，兼容性极佳）
android.api = 33

# (int) 最低支持的安卓系统版本（21 对应 Android 5.0，基本覆盖全网所有老手机）
android.minapi = 21

# (list) 打包的 CPU 架构，下面这两个组合可以完美兼容目前市面上 99% 的安卓手机
android.archs = armeabi-v7a, arm64-v8a

# (bool) 开启现代安卓依赖库支持（必须为 True，否则打包容易报错）
android.enable_androidx = True

# (str) 或者是自动给打包的 APK 加上一些安全标签
android.skip_apk_rescale = 1


[buildozer]

# (int) 日志输出级别 (2 代表 Debug，能让你在 GitHub Actions 里看到最详细的打包进度)
log_level = 2

# (int) 是否在 root 权限警告时中断
warn_on_root = 1
