@echo off
chcp 65001 >nul
cd /d %~dp0
title PatentKB Sync
python sync.py loop
