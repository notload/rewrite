[app]
title = 听写助手
package.name = dictation
package.domain = org.example

source.dir = .
source.include_exts = py,png,jpg,kv,atlas

version = 1.0

requirements = python3,kivy,openpyxl,requests

orientation = portrait

# 使用 GitHub Actions 预装的 SDK/NDK（API 34 是镜像里已有的）
android.sdk_path = /usr/local/lib/android/sdk
android.ndk_path = /usr/local/lib/android/sdk/ndk/27.3.13750724
android.api = 34
android.minapi = 21
android.accept_sdk_license = True
android.skip_update = True

# 明确指定 bootstrap（默认 sdl2）
bootstrap = sdl2

# 权限（音频需要 Internet，存储需要读写）
android.permissions = INTERNET,READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE

[buildozer]
log_level = 2
warn_on_root = 1
