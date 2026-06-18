---
title: モジュール一覧
description: VBA モジュール（.bas / .cls）のコピペ用ソースコード一覧
---

# モジュール一覧

knowledgevba を構成する VBA モジュールのソースコード一覧です。各ページに該当ファイルの全文がコピペできる形で掲載されています。

[インストール手順](../install.md) の **STEP 6** で、これらを `installer\vba_modules\<役割>\` 配下に保存してください。

**更新日** は canonical ソースの最終更新日時です。手元のファイルより新しければ差し替えてください。

!!! info "保存時のお願い"
    - メモ帳で **[名前を付けて保存]** → 文字コードは **ANSI**（Shift-JIS）を選んでください
    - UTF-8 で保存すると日本語が文字化けして動かなくなります
    - ファイル名は **大文字小文字を区別しません**（VBE 側で正しい名前が付きます）

---

## 管理.xlsm 用 (`installer\vba_modules\admin\`)

**9 ファイル**

| ファイル名 | 種類 | 更新日 |
|---|---|---|
| [`ThisWorkbook.cls`](admin/thisworkbook.cls.md) | クラス | 2026-06-11 19:15 JST |
| [`clsFormatDesignScreen.cls`](admin/clsformatdesignscreen.cls.md) | クラス | 2026-06-07 06:30 JST |
| [`clsFormatListScreen.cls`](admin/clsformatlistscreen.cls.md) | クラス | 2026-06-04 12:30 JST |
| [`clsFormatPreviewScreen.cls`](admin/clsformatpreviewscreen.cls.md) | クラス | 2026-06-06 22:08 JST |
| [`clsMigrationScreen.cls`](admin/clsmigrationscreen.cls.md) | クラス | 2026-06-04 12:30 JST |
| [`clsStorageScreen.cls`](admin/clsstoragescreen.cls.md) | クラス | 2026-06-03 23:22 JST |
| [`clsSystemSettingsScreen.cls`](admin/clssystemsettingsscreen.cls.md) | クラス | 2026-06-04 12:30 JST |
| [`modEntryFormat.bas`](admin/modentryformat.bas.md) | 標準 | 2026-06-13 01:10 JST |
| [`modEntrySettings.bas`](admin/modentrysettings.bas.md) | 標準 | 2026-06-17 22:19 JST |

## 検索.xlsm 用 (`installer\vba_modules\search\`)

**8 ファイル**

| ファイル名 | 種類 | 更新日 |
|---|---|---|
| [`ThisWorkbook.cls`](search/thisworkbook.cls.md) | クラス | 2026-06-11 16:37 JST |
| [`clsKnowledgeEditScreen.cls`](search/clsknowledgeeditscreen.cls.md) | クラス | 2026-06-09 22:15 JST |
| [`clsKnowledgeRegisterScreen.cls`](search/clsknowledgeregisterscreen.cls.md) | クラス | 2026-06-09 22:15 JST |
| [`clsKnowledgeViewScreen.cls`](search/clsknowledgeviewscreen.cls.md) | クラス | 2026-06-03 23:22 JST |
| [`clsSearchScreen.cls`](search/clssearchscreen.cls.md) | クラス | 2026-06-03 23:22 JST |
| [`modEntryKnowledge.bas`](search/modentryknowledge.bas.md) | 標準 | 2026-06-09 22:15 JST |
| [`modEntrySearch.bas`](search/modentrysearch.bas.md) | 標準 | 2026-06-12 22:19 JST |
| [`modKnowledgeEntryHelpers.bas`](search/modknowledgeentryhelpers.bas.md) | 標準 | 2026-06-09 22:15 JST |

## 共通モジュール (`installer\vba_modules\common\`)

**51 ファイル**

| ファイル名 | 種類 | 更新日 |
|---|---|---|
| [`ClsStanzaSection.cls`](common/clsstanzasection.cls.md) | クラス | 2026-06-03 23:22 JST |
| [`ClsStanzaValidationIssue.cls`](common/clsstanzavalidationissue.cls.md) | クラス | 2026-06-03 23:22 JST |
| [`ClsStanzaValidationResult.cls`](common/clsstanzavalidationresult.cls.md) | クラス | 2026-06-03 23:22 JST |
| [`IScreenRenderer.cls`](common/iscreenrenderer.cls.md) | クラス | 2026-06-03 23:22 JST |
| [`clsButtonSpec.cls`](common/clsbuttonspec.cls.md) | クラス | 2026-06-03 23:22 JST |
| [`clsCellAddrHelper.cls`](common/clscelladdrhelper.cls.md) | クラス | 2026-06-04 12:30 JST |
| [`clsCellBinding.cls`](common/clscellbinding.cls.md) | クラス | 2026-06-03 23:22 JST |
| [`clsCellIO.cls`](common/clscellio.cls.md) | クラス | 2026-06-03 23:22 JST |
| [`clsControlSpec.cls`](common/clscontrolspec.cls.md) | クラス | 2026-06-03 23:22 JST |
| [`clsFieldMigrator.cls`](common/clsfieldmigrator.cls.md) | クラス | 2026-06-03 23:22 JST |
| [`clsFieldSpec.cls`](common/clsfieldspec.cls.md) | クラス | 2026-06-06 22:47 JST |
| [`clsFormSpec.cls`](common/clsformspec.cls.md) | クラス | 2026-06-03 23:22 JST |
| [`clsFormatManager.cls`](common/clsformatmanager.cls.md) | クラス | 2026-06-05 01:27 JST |
| [`clsGridIO.cls`](common/clsgridio.cls.md) | クラス | 2026-06-11 15:18 JST |
| [`clsKnowledgeManager.cls`](common/clsknowledgemanager.cls.md) | クラス | 2026-06-05 19:07 JST |
| [`clsLogEntry.cls`](common/clslogentry.cls.md) | クラス | 2026-06-03 23:22 JST |
| [`clsLogger.cls`](common/clslogger.cls.md) | クラス | 2026-06-12 01:23 JST |
| [`clsScreenSpec.cls`](common/clsscreenspec.cls.md) | クラス | 2026-06-05 01:27 JST |
| [`clsSearchEngine.cls`](common/clssearchengine.cls.md) | クラス | 2026-06-05 19:06 JST |
| [`clsSectionSpec.cls`](common/clssectionspec.cls.md) | クラス | 2026-06-03 23:22 JST |
| [`clsSetupOrchestrator.cls`](common/clssetuporchestrator.cls.md) | クラス | 2026-06-17 17:29 JST |
| [`clsSheetRenderer.cls`](common/clssheetrenderer.cls.md) | クラス | 2026-06-14 14:15 JST |
| [`clsStorageResolver.cls`](common/clsstorageresolver.cls.md) | クラス | 2026-06-05 01:27 JST |
| [`clsUserFormRenderer.cls`](common/clsuserformrenderer.cls.md) | クラス | 2026-06-12 08:51 JST |
| [`modBtnGuard.bas`](common/modbtnguard.bas.md) | 標準 | 2026-06-07 07:38 JST |
| [`modBtnMessages.bas`](common/modbtnmessages.bas.md) | 標準 | 2026-06-05 01:27 JST |
| [`modButtonWiring.bas`](common/modbuttonwiring.bas.md) | 標準 | 2026-06-03 23:22 JST |
| [`modCommon.bas`](common/modcommon.bas.md) | 標準 | 2026-06-12 01:22 JST |
| [`modConfigHolder.bas`](common/modconfigholder.bas.md) | 標準 | 2026-06-17 22:19 JST |
| [`modConfigLoader.bas`](common/modconfigloader.bas.md) | 標準 | 2026-06-17 22:19 JST |
| [`modDateUtil.bas`](common/moddateutil.bas.md) | 標準 | 2026-06-03 23:22 JST |
| [`modEntryUserForm.bas`](common/modentryuserform.bas.md) | 標準 | 2026-06-17 01:05 JST |
| [`modFactory.bas`](common/modfactory.bas.md) | 標準 | 2026-06-04 12:30 JST |
| [`modFileIO.bas`](common/modfileio.bas.md) | 標準 | 2026-06-03 23:22 JST |
| [`modFormBuilder.bas`](common/modformbuilder.bas.md) | 標準 | 2026-06-03 23:22 JST |
| [`modFormControlWiring.bas`](common/modformcontrolwiring.bas.md) | 標準 | 2026-06-04 12:30 JST |
| [`modFormatColumns.bas`](common/modformatcolumns.bas.md) | 標準 | 2026-06-03 23:22 JST |
| [`modFormatLoader.bas`](common/modformatloader.bas.md) | 標準 | 2026-06-09 22:38 JST |
| [`modInstallConfig.bas`](common/modinstallconfig.bas.md) | 標準 | 2026-06-18 11:06 JST |
| [`modKnowledgeFileIO.bas`](common/modknowledgefileio.bas.md) | 標準 | 2026-06-11 15:52 JST |
| [`modLogIds.bas`](common/modlogids.bas.md) | 標準 | 2026-06-03 23:22 JST |
| [`modPreviewForm.bas`](common/modpreviewform.bas.md) | 標準 | 2026-06-04 12:30 JST |
| [`modRefresh.bas`](common/modrefresh.bas.md) | 標準 | 2026-06-12 22:33 JST |
| [`modScreenRender.bas`](common/modscreenrender.bas.md) | 標準 | 2026-06-03 23:22 JST |
| [`modSetup.bas`](common/modsetup.bas.md) | 標準 | 2026-06-10 01:26 JST |
| [`modSheetButtons.bas`](common/modsheetbuttons.bas.md) | 標準 | 2026-06-12 01:23 JST |
| [`modSheetMap.bas`](common/modsheetmap.bas.md) | 標準 | 2026-06-03 23:22 JST |
| [`modStanzaIO.bas`](common/modstanzaio.bas.md) | 標準 | 2026-06-05 01:27 JST |
| [`modStringUtil.bas`](common/modstringutil.bas.md) | 標準 | 2026-06-05 01:27 JST |
| [`modUILoader.bas`](common/moduiloader.bas.md) | 標準 | 2026-06-12 23:05 JST |
| [`modUserFormCallback.bas`](common/moduserformcallback.bas.md) | 標準 | 2026-06-12 08:51 JST |

