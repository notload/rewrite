[app]
title = 听写助手
package.name = dictation
package.domain = org.example

source.dir = .
source.include_exts = py,png,jpg,kv,atlas

version = 1.0

requirements = python3,kivy,openpyxl,requests

orientation = portrait

# 使用 GitHub Actions 预装的 SDK
android.sdk_path = /usr/local/lib/android/sdk
android.ndk_path = /usr/local/lib/android/sdk/ndk/27.3.13750724
android.api = 30
android.minapi = 21
android.accept_sdk_license = True
android.skip_update = True

# 权限
android.permissions = INTERNET,READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE

[buildozer]
log_level = 2
warn_on_root = 1
