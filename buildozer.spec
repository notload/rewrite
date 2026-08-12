[app]
title = 听写助手
package.name = dictation
package.domain = org.example

source.dir = .
source.include_exts = py,png,jpg,kv,atlas

version = 1.0

requirements = python3,kivy,openpyxl,requests

orientation = portrait

osx.python_version = 3
osx.kivy_version = 2.1.0

# 安卓权限
android.permissions = INTERNET,READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE
android.api = 31
android.minapi = 21
android.sdk_path = ~/Android/Sdk   # 如果自动下载失败，可手动指定
android.ndk_path = ~/Android/android-ndk-r25c
