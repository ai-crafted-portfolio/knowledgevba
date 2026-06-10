---
title: モジュール一覧
description: VBA モジュール（.bas / .cls）のコピペ用ソースコード一覧
---

# モジュール一覧

knowledgevba を構成する VBA モジュールのソースコード一覧です。各ページに該当ファイルの全文がコピペできる形で掲載されています。

[インストール手順](../install.md) の **STEP 6** で、これらを `installer\vba_modules\<役割>\` 配下に保存してください。

!!! info "保存時のお願い"
    - メモ帳で **[名前を付けて保存]** → 文字コードは **ANSI**（Shift-JIS）を選んでください
    - UTF-8 で保存すると日本語が文字化けして動かなくなります
    - ファイル名は **大文字小文字を区別しません**（VBE 側で正しい名前が付きます）

---

## 管理.xlsm 用 (`installer\vba_modules\admin\`)

**10 ファイル**

| ファイル名 | 種類 |
|---|---|
| [`clsFormatDesignScreen.cls`](admin/clsformatdesignscreen.cls.md) | クラス |
| [`clsFormatListScreen.cls`](admin/clsformatlistscreen.cls.md) | クラス |
| [`clsFormatPreviewScreen.cls`](admin/clsformatpreviewscreen.cls.md) | クラス |
| [`clsLogScreen.cls`](admin/clslogscreen.cls.md) | クラス |
| [`clsMigrationScreen.cls`](admin/clsmigrationscreen.cls.md) | クラス |
| [`clsStorageScreen.cls`](admin/clsstoragescreen.cls.md) | クラス |
| [`clsSystemSettingsScreen.cls`](admin/clssystemsettingsscreen.cls.md) | クラス |
| [`modEntryFormat.bas`](admin/modentryformat.bas.md) | 標準 |
| [`modEntrySettings.bas`](admin/modentrysettings.bas.md) | 標準 |
| [`ThisWorkbook.cls`](admin/thisworkbook.cls.md) | クラス |

## 登録修正.xlsm 用 (`installer\vba_modules\register\`)

**5 ファイル**

| ファイル名 | 種類 |
|---|---|
| [`clsKnowledgeEditScreen.cls`](register/clsknowledgeeditscreen.cls.md) | クラス |
| [`clsKnowledgeRegisterScreen.cls`](register/clsknowledgeregisterscreen.cls.md) | クラス |
| [`modEntryKnowledge.bas`](register/modentryknowledge.bas.md) | 標準 |
| [`modKnowledgeEntryHelpers.bas`](register/modknowledgeentryhelpers.bas.md) | 標準 |
| [`ThisWorkbook.cls`](register/thisworkbook.cls.md) | クラス |

## 検索.xlsm 用 (`installer\vba_modules\search\`)

**4 ファイル**

| ファイル名 | 種類 |
|---|---|
| [`clsKnowledgeViewScreen.cls`](search/clsknowledgeviewscreen.cls.md) | クラス |
| [`clsSearchScreen.cls`](search/clssearchscreen.cls.md) | クラス |
| [`modEntrySearch.bas`](search/modentrysearch.bas.md) | 標準 |
| [`ThisWorkbook.cls`](search/thisworkbook.cls.md) | クラス |

## 共通 (`installer\vba_modules\common\`)

**64 ファイル**

| ファイル名 | 種類 |
|---|---|
| [`clsButtonSpec.cls`](common/clsbuttonspec.cls.md) | クラス |
| [`clsCellAddrHelper.cls`](common/clscelladdrhelper.cls.md) | クラス |
| [`clsCellBinding.cls`](common/clscellbinding.cls.md) | クラス |
| [`clsCellIO.cls`](common/clscellio.cls.md) | クラス |
| [`clsControlSpec.cls`](common/clscontrolspec.cls.md) | クラス |
| [`clsFieldMigrator.cls`](common/clsfieldmigrator.cls.md) | クラス |
| [`clsFieldSpec.cls`](common/clsfieldspec.cls.md) | クラス |
| [`clsFormatManager.cls`](common/clsformatmanager.cls.md) | クラス |
| [`clsFormSpec.cls`](common/clsformspec.cls.md) | クラス |
| [`clsGridIO.cls`](common/clsgridio.cls.md) | クラス |
| [`clsKnowledgeManager.cls`](common/clsknowledgemanager.cls.md) | クラス |
| [`clsLogEntry.cls`](common/clslogentry.cls.md) | クラス |
| [`clsLogger.cls`](common/clslogger.cls.md) | クラス |
| [`clsScreenSpec.cls`](common/clsscreenspec.cls.md) | クラス |
| [`clsSearchEngine.cls`](common/clssearchengine.cls.md) | クラス |
| [`clsSectionSpec.cls`](common/clssectionspec.cls.md) | クラス |
| [`clsSetupOrchestrator.cls`](common/clssetuporchestrator.cls.md) | クラス |
| [`clsSheetRenderer.cls`](common/clssheetrenderer.cls.md) | クラス |
| [`ClsStanzaSection.cls`](common/clsstanzasection.cls.md) | クラス |
| [`ClsStanzaValidationIssue.cls`](common/clsstanzavalidationissue.cls.md) | クラス |
| [`ClsStanzaValidationResult.cls`](common/clsstanzavalidationresult.cls.md) | クラス |
| [`clsStorageResolver.cls`](common/clsstorageresolver.cls.md) | クラス |
| [`clsUserFormRenderer.cls`](common/clsuserformrenderer.cls.md) | クラス |
| [`IScreenRenderer.cls`](common/iscreenrenderer.cls.md) | クラス |
| [`modBtnGuard.bas`](common/modbtnguard.bas.md) | 標準 |
| [`modBtnMessages.bas`](common/modbtnmessages.bas.md) | 標準 |
| [`modButtonWiring.bas`](common/modbuttonwiring.bas.md) | 標準 |
| [`modCommon.bas`](common/modcommon.bas.md) | 標準 |
| [`modConfigHolder.bas`](common/modconfigholder.bas.md) | 標準 |
| [`modConfigLoader.bas`](common/modconfigloader.bas.md) | 標準 |
| [`modDateUtil.bas`](common/moddateutil.bas.md) | 標準 |
| [`modE2E_Abnormal.bas`](common/mode2e_abnormal.bas.md) | 標準 |
| [`modE2E_ADRCompliance.bas`](common/mode2e_adrcompliance.bas.md) | 標準 |
| [`modE2E_AllButtons.bas`](common/mode2e_allbuttons.bas.md) | 標準 |
| [`modE2E_Backup.bas`](common/mode2e_backup.bas.md) | 標準 |
| [`modE2E_DataFormat.bas`](common/mode2e_dataformat.bas.md) | 標準 |
| [`modE2E_Environment.bas`](common/mode2e_environment.bas.md) | 標準 |
| [`modE2E_FieldType.bas`](common/mode2e_fieldtype.bas.md) | 標準 |
| [`modE2E_Lifecycle.bas`](common/mode2e_lifecycle.bas.md) | 標準 |
| [`modE2E_MultiRole.bas`](common/mode2e_multirole.bas.md) | 標準 |
| [`modE2E_Search.bas`](common/mode2e_search.bas.md) | 標準 |
| [`modE2E_UserForm.bas`](common/mode2e_userform.bas.md) | 標準 |
| [`modEntryUserForm.bas`](common/modentryuserform.bas.md) | 標準 |
| [`modFactory.bas`](common/modfactory.bas.md) | 標準 |
| [`modFileIO.bas`](common/modfileio.bas.md) | 標準 |
| [`modFormatColumns.bas`](common/modformatcolumns.bas.md) | 標準 |
| [`modFormatLoader.bas`](common/modformatloader.bas.md) | 標準 |
| [`modFormBuilder.bas`](common/modformbuilder.bas.md) | 標準 |
| [`modFormControlWiring.bas`](common/modformcontrolwiring.bas.md) | 標準 |
| [`modKnowledgeFileIO.bas`](common/modknowledgefileio.bas.md) | 標準 |
| [`modLogIds.bas`](common/modlogids.bas.md) | 標準 |
| [`modPerfTest.bas`](common/modperftest.bas.md) | 標準 |
| [`modPreviewForm.bas`](common/modpreviewform.bas.md) | 標準 |
| [`modRefresh.bas`](common/modrefresh.bas.md) | 標準 |
| [`modScreenRender.bas`](common/modscreenrender.bas.md) | 標準 |
| [`modSetup.bas`](common/modsetup.bas.md) | 標準 |
| [`modSheetButtons.bas`](common/modsheetbuttons.bas.md) | 標準 |
| [`modSheetMap.bas`](common/modsheetmap.bas.md) | 標準 |
| [`modSpecExamples.bas`](common/modspecexamples.bas.md) | 標準 |
| [`modStanzaIO.bas`](common/modstanzaio.bas.md) | 標準 |
| [`modStringUtil.bas`](common/modstringutil.bas.md) | 標準 |
| [`modTestHelpers.bas`](common/modtesthelpers.bas.md) | 標準 |
| [`modUILoader.bas`](common/moduiloader.bas.md) | 標準 |
| [`modUserFormCallback.bas`](common/moduserformcallback.bas.md) | 標準 |

---

合計 **83 ファイル**
