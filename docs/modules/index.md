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

**10 ファイル**

| ファイル名 | 種類 | 更新日 |
|---|---|---|
| [`ThisWorkbook.cls`](admin/thisworkbook.cls.md) | クラス | 2026-06-11 19:15 |
| [`clsFormatDesignScreen.cls`](admin/clsformatdesignscreen.cls.md) | クラス | 2026-06-07 06:30 |
| [`clsFormatListScreen.cls`](admin/clsformatlistscreen.cls.md) | クラス | 2026-06-04 12:30 |
| [`clsFormatPreviewScreen.cls`](admin/clsformatpreviewscreen.cls.md) | クラス | 2026-06-06 22:08 |
| [`clsLogScreen.cls`](admin/clslogscreen.cls.md) | クラス | 2026-06-04 12:30 |
| [`clsMigrationScreen.cls`](admin/clsmigrationscreen.cls.md) | クラス | 2026-06-04 12:30 |
| [`clsStorageScreen.cls`](admin/clsstoragescreen.cls.md) | クラス | 2026-06-03 23:22 |
| [`clsSystemSettingsScreen.cls`](admin/clssystemsettingsscreen.cls.md) | クラス | 2026-06-04 12:30 |
| [`modEntryFormat.bas`](admin/modentryformat.bas.md) | 標準 | 2026-06-12 01:22 |
| [`modEntrySettings.bas`](admin/modentrysettings.bas.md) | 標準 | 2026-06-11 22:50 |

## 登録修正.xlsm 用 (`installer\vba_modules\register\`) ※v2.3 では廃止ブック（参考掲載）

**5 ファイル**

| ファイル名 | 種類 | 更新日 |
|---|---|---|
| [`ThisWorkbook.cls`](register/thisworkbook.cls.md) | クラス | 2026-06-06 08:22 |
| [`clsKnowledgeEditScreen.cls`](register/clsknowledgeeditscreen.cls.md) | クラス | 2026-06-03 23:22 |
| [`clsKnowledgeRegisterScreen.cls`](register/clsknowledgeregisterscreen.cls.md) | クラス | 2026-06-03 23:22 |
| [`modEntryKnowledge.bas`](register/modentryknowledge.bas.md) | 標準 | 2026-06-04 21:35 |
| [`modKnowledgeEntryHelpers.bas`](register/modknowledgeentryhelpers.bas.md) | 標準 | 2026-06-03 23:22 |

## 検索.xlsm 用 (`installer\vba_modules\search\`)

**8 ファイル**

| ファイル名 | 種類 | 更新日 |
|---|---|---|
| [`ThisWorkbook.cls`](search/thisworkbook.cls.md) | クラス | 2026-06-11 16:37 |
| [`clsKnowledgeEditScreen.cls`](search/clsknowledgeeditscreen.cls.md) | クラス | 2026-06-09 22:15 |
| [`clsKnowledgeRegisterScreen.cls`](search/clsknowledgeregisterscreen.cls.md) | クラス | 2026-06-09 22:15 |
| [`clsKnowledgeViewScreen.cls`](search/clsknowledgeviewscreen.cls.md) | クラス | 2026-06-03 23:22 |
| [`clsSearchScreen.cls`](search/clssearchscreen.cls.md) | クラス | 2026-06-03 23:22 |
| [`modEntryKnowledge.bas`](search/modentryknowledge.bas.md) | 標準 | 2026-06-09 22:15 |
| [`modEntrySearch.bas`](search/modentrysearch.bas.md) | 標準 | 2026-06-12 01:22 |
| [`modKnowledgeEntryHelpers.bas`](search/modknowledgeentryhelpers.bas.md) | 標準 | 2026-06-09 22:15 |

## 共通モジュール (`installer\vba_modules\common\`)

**64 ファイル**

| ファイル名 | 種類 | 更新日 |
|---|---|---|
| [`ClsStanzaSection.cls`](common/clsstanzasection.cls.md) | クラス | 2026-06-03 23:22 |
| [`ClsStanzaValidationIssue.cls`](common/clsstanzavalidationissue.cls.md) | クラス | 2026-06-03 23:22 |
| [`ClsStanzaValidationResult.cls`](common/clsstanzavalidationresult.cls.md) | クラス | 2026-06-03 23:22 |
| [`IScreenRenderer.cls`](common/iscreenrenderer.cls.md) | クラス | 2026-06-03 23:22 |
| [`clsButtonSpec.cls`](common/clsbuttonspec.cls.md) | クラス | 2026-06-03 23:22 |
| [`clsCellAddrHelper.cls`](common/clscelladdrhelper.cls.md) | クラス | 2026-06-04 12:30 |
| [`clsCellBinding.cls`](common/clscellbinding.cls.md) | クラス | 2026-06-03 23:22 |
| [`clsCellIO.cls`](common/clscellio.cls.md) | クラス | 2026-06-03 23:22 |
| [`clsControlSpec.cls`](common/clscontrolspec.cls.md) | クラス | 2026-06-03 23:22 |
| [`clsFieldMigrator.cls`](common/clsfieldmigrator.cls.md) | クラス | 2026-06-03 23:22 |
| [`clsFieldSpec.cls`](common/clsfieldspec.cls.md) | クラス | 2026-06-06 22:47 |
| [`clsFormSpec.cls`](common/clsformspec.cls.md) | クラス | 2026-06-03 23:22 |
| [`clsFormatManager.cls`](common/clsformatmanager.cls.md) | クラス | 2026-06-05 01:27 |
| [`clsGridIO.cls`](common/clsgridio.cls.md) | クラス | 2026-06-11 15:18 |
| [`clsKnowledgeManager.cls`](common/clsknowledgemanager.cls.md) | クラス | 2026-06-05 19:07 |
| [`clsLogEntry.cls`](common/clslogentry.cls.md) | クラス | 2026-06-03 23:22 |
| [`clsLogger.cls`](common/clslogger.cls.md) | クラス | 2026-06-12 01:23 |
| [`clsScreenSpec.cls`](common/clsscreenspec.cls.md) | クラス | 2026-06-05 01:27 |
| [`clsSearchEngine.cls`](common/clssearchengine.cls.md) | クラス | 2026-06-05 19:06 |
| [`clsSectionSpec.cls`](common/clssectionspec.cls.md) | クラス | 2026-06-03 23:22 |
| [`clsSetupOrchestrator.cls`](common/clssetuporchestrator.cls.md) | クラス | 2026-06-09 22:14 |
| [`clsSheetRenderer.cls`](common/clssheetrenderer.cls.md) | クラス | 2026-06-05 22:31 |
| [`clsStorageResolver.cls`](common/clsstorageresolver.cls.md) | クラス | 2026-06-05 01:27 |
| [`clsUserFormRenderer.cls`](common/clsuserformrenderer.cls.md) | クラス | 2026-06-12 08:51 |
| [`modBtnGuard.bas`](common/modbtnguard.bas.md) | 標準 | 2026-06-07 07:38 |
| [`modBtnMessages.bas`](common/modbtnmessages.bas.md) | 標準 | 2026-06-05 01:27 |
| [`modButtonWiring.bas`](common/modbuttonwiring.bas.md) | 標準 | 2026-06-03 23:22 |
| [`modCommon.bas`](common/modcommon.bas.md) | 標準 | 2026-06-12 01:22 |
| [`modConfigHolder.bas`](common/modconfigholder.bas.md) | 標準 | 2026-06-03 23:22 |
| [`modConfigLoader.bas`](common/modconfigloader.bas.md) | 標準 | 2026-06-07 00:50 |
| [`modDateUtil.bas`](common/moddateutil.bas.md) | 標準 | 2026-06-03 23:22 |
| [`modE2E_ADRCompliance.bas`](common/mode2e_adrcompliance.bas.md) | 標準 | 2026-06-08 12:53 |
| [`modE2E_Abnormal.bas`](common/mode2e_abnormal.bas.md) | 標準 | 2026-06-08 12:53 |
| [`modE2E_AllButtons.bas`](common/mode2e_allbuttons.bas.md) | 標準 | 2026-06-09 09:23 |
| [`modE2E_Backup.bas`](common/mode2e_backup.bas.md) | 標準 | 2026-06-08 12:53 |
| [`modE2E_DataFormat.bas`](common/mode2e_dataformat.bas.md) | 標準 | 2026-06-08 12:53 |
| [`modE2E_Environment.bas`](common/mode2e_environment.bas.md) | 標準 | 2026-06-08 12:53 |
| [`modE2E_FieldType.bas`](common/mode2e_fieldtype.bas.md) | 標準 | 2026-06-08 12:53 |
| [`modE2E_Lifecycle.bas`](common/mode2e_lifecycle.bas.md) | 標準 | 2026-06-08 12:53 |
| [`modE2E_MultiRole.bas`](common/mode2e_multirole.bas.md) | 標準 | 2026-06-08 12:53 |
| [`modE2E_Search.bas`](common/mode2e_search.bas.md) | 標準 | 2026-06-08 12:53 |
| [`modE2E_UserForm.bas`](common/mode2e_userform.bas.md) | 標準 | 2026-06-08 12:53 |
| [`modEntryUserForm.bas`](common/modentryuserform.bas.md) | 標準 | 2026-06-04 12:30 |
| [`modFactory.bas`](common/modfactory.bas.md) | 標準 | 2026-06-04 12:30 |
| [`modFileIO.bas`](common/modfileio.bas.md) | 標準 | 2026-06-03 23:22 |
| [`modFormBuilder.bas`](common/modformbuilder.bas.md) | 標準 | 2026-06-03 23:22 |
| [`modFormControlWiring.bas`](common/modformcontrolwiring.bas.md) | 標準 | 2026-06-04 12:30 |
| [`modFormatColumns.bas`](common/modformatcolumns.bas.md) | 標準 | 2026-06-03 23:22 |
| [`modFormatLoader.bas`](common/modformatloader.bas.md) | 標準 | 2026-06-09 22:38 |
| [`modKnowledgeFileIO.bas`](common/modknowledgefileio.bas.md) | 標準 | 2026-06-11 15:52 |
| [`modLogIds.bas`](common/modlogids.bas.md) | 標準 | 2026-06-03 23:22 |
| [`modPerfTest.bas`](common/modperftest.bas.md) | 標準 | 2026-06-05 17:15 |
| [`modPreviewForm.bas`](common/modpreviewform.bas.md) | 標準 | 2026-06-04 12:30 |
| [`modRefresh.bas`](common/modrefresh.bas.md) | 標準 | 2026-06-09 09:22 |
| [`modScreenRender.bas`](common/modscreenrender.bas.md) | 標準 | 2026-06-03 23:22 |
| [`modSetup.bas`](common/modsetup.bas.md) | 標準 | 2026-06-10 01:26 |
| [`modSheetButtons.bas`](common/modsheetbuttons.bas.md) | 標準 | 2026-06-12 01:23 |
| [`modSheetMap.bas`](common/modsheetmap.bas.md) | 標準 | 2026-06-03 23:22 |
| [`modSpecExamples.bas`](common/modspecexamples.bas.md) | 標準 | 2026-06-05 01:27 |
| [`modStanzaIO.bas`](common/modstanzaio.bas.md) | 標準 | 2026-06-05 01:27 |
| [`modStringUtil.bas`](common/modstringutil.bas.md) | 標準 | 2026-06-05 01:27 |
| [`modTestHelpers.bas`](common/modtesthelpers.bas.md) | 標準 | 2026-06-03 23:22 |
| [`modUILoader.bas`](common/moduiloader.bas.md) | 標準 | 2026-06-12 01:22 |
| [`modUserFormCallback.bas`](common/moduserformcallback.bas.md) | 標準 | 2026-06-12 08:51 |

