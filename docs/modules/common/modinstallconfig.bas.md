---
title: modInstallConfig.bas
description: modInstallConfig.bas のソースコード（コピペ用）
---

# modInstallConfig.bas

**配置先**: 共通モジュール（検索.xlsm / 管理.xlsm 共通）
**種類**: 標準モジュール
**更新日**: 2026-06-18 11:06 JST

---

## 概要

このモジュールは、データやログなどの保存先フォルダを決める **パス定数** をまとめたものです。

通常はインストーラ（`_auto_install.ps1`）が導入時に自動生成して取り込むため、手作業での用意は不要です。インストーラを使わず手動で導入したい場合や、内容を確認したい場合のために、このページに全文を掲載しています。

保存先を既定（`C:\KnowledgeMgr\...`）から変更したい場合は、下記 6 行の各パスを希望のフォルダに書き換えてから取り込んでください。

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `modInstallConfig.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS / CP932）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

VBA エディタで取り込むときは、**ファイル → ファイルのインポート** から保存した `modInstallConfig.bas` を選びます。

---

## ソースコード

```vb
Attribute VB_Name = "modInstallConfig"
' ============================================================
' modInstallConfig
' データ・フォーマット・UI・バックアップ・ログなどの
' 既定の保存先フォルダを定義するパス定数モジュール。
' 通常はインストーラが導入時に自動生成して取り込みます。
' ============================================================
Option Explicit
Public Const DEFAULT_DATA_DIR    As String = "C:\KnowledgeMgr\data\"
Public Const DEFAULT_FORMAT_DIR  As String = "C:\KnowledgeMgr\formats\"
Public Const DEFAULT_UI_DIR      As String = "C:\KnowledgeMgr\ui\"
Public Const DEFAULT_BACKUP_DIR  As String = "C:\KnowledgeMgr\backup\"
Public Const DEFAULT_LOG_DIR     As String = "C:\KnowledgeMgr\log\"
Public Const DEFAULT_CONFIG_DIR  As String = "C:\KnowledgeMgr\"
```
