---
title: モジュール一覧
description: VBA モジュール（.bas / .cls）のコピペ用ソースコード一覧
---

# モジュール一覧

knowledgevba を構成する VBA モジュールのソースコード一覧です。各ページに該当ファイルの全文がコピペできる形で掲載されています。

[インストール手順](../install.md) の **STEP 6** で、これらを `installer\vba_modules\<役割>\` 配下に保存してください。

**更新日** はソースの最終更新日時です。手元のファイルより新しければ差し替えてください。

!!! info "保存時のお願い"
    - メモ帳で **[名前を付けて保存]** → 文字コードは **ANSI**（Shift-JIS）を選んでください
    - UTF-8 で保存すると日本語が文字化けして動かなくなります
    - ファイル名は **大文字小文字を区別しません**（VBE 側で正しい名前が付きます）

---

## 管理.xlsm 用 (`installer\vba_modules\admin\`)

**8 ファイル**

| ファイル名 | 種類 | 更新日 |
|---|---|---|
| [`ThisWorkbook.cls`](admin/thisworkbook.cls.md) | クラス | 2026-06-30 14:44 JST |
| [`clsFormatDesignScreen.cls`](admin/clsformatdesignscreen.cls.md) | クラス | 2026-06-30 14:44 JST |
| [`clsFormatListScreen.cls`](admin/clsformatlistscreen.cls.md) | クラス | 2026-06-30 14:44 JST |
| [`clsFormatPreviewScreen.cls`](admin/clsformatpreviewscreen.cls.md) | クラス | 2026-06-30 14:44 JST |
| [`clsMigrationScreen.cls`](admin/clsmigrationscreen.cls.md) | クラス | 2026-06-30 14:44 JST |
| [`clsStorageScreen.cls`](admin/clsstoragescreen.cls.md) | クラス | 2026-06-30 14:44 JST |
| [`modEntryFormat.bas`](admin/modentryformat.bas.md) | 標準 | 2026-06-30 14:44 JST |
| [`modEntrySettings.bas`](admin/modentrysettings.bas.md) | 標準 | 2026-06-30 14:44 JST |

## 検索.xlsm 用 (`installer\vba_modules\search\`)

**8 ファイル**

| ファイル名 | 種類 | 更新日 |
|---|---|---|
| [`ThisWorkbook.cls`](search/thisworkbook.cls.md) | クラス | 2026-06-30 14:44 JST |
| [`clsKnowledgeEditScreen.cls`](search/clsknowledgeeditscreen.cls.md) | クラス | 2026-06-30 14:44 JST |
| [`clsKnowledgeRegisterScreen.cls`](search/clsknowledgeregisterscreen.cls.md) | クラス | 2026-06-30 14:44 JST |
| [`clsKnowledgeViewScreen.cls`](search/clsknowledgeviewscreen.cls.md) | クラス | 2026-06-30 14:44 JST |
| [`clsSearchScreen.cls`](search/clssearchscreen.cls.md) | クラス | 2026-06-30 14:44 JST |
| [`modEntryKnowledge.bas`](search/modentryknowledge.bas.md) | 標準 | 2026-06-30 14:44 JST |
| [`modEntrySearch.bas`](search/modentrysearch.bas.md) | 標準 | 2026-06-30 14:44 JST |
| [`modKnowledgeEntryHelpers.bas`](search/modknowledgeentryhelpers.bas.md) | 標準 | 2026-06-30 14:44 JST |

## 共通モジュール (`installer\vba_modules\common\`)

**50 ファイル**

| ファイル名 | 種類 | 更新日 |
|---|---|---|
| [`ClsStanzaSection.cls`](common/clsstanzasection.cls.md) | クラス | 2026-06-30 14:44 JST |
| [`ClsStanzaValidationIssue.cls`](common/clsstanzavalidationissue.cls.md) | クラス | 2026-06-30 14:44 JST |
| [`ClsStanzaValidationResult.cls`](common/clsstanzavalidationresult.cls.md) | クラス | 2026-06-30 14:44 JST |
| [`IScreenRenderer.cls`](common/iscreenrenderer.cls.md) | クラス | 2026-06-30 14:44 JST |
| [`clsButtonSpec.cls`](common/clsbuttonspec.cls.md) | クラス | 2026-06-30 14:44 JST |
| [`clsCellAddrHelper.cls`](common/clscelladdrhelper.cls.md) | クラス | 2026-06-30 14:44 JST |
| [`clsCellBinding.cls`](common/clscellbinding.cls.md) | クラス | 2026-06-30 14:44 JST |
| [`clsCellIO.cls`](common/clscellio.cls.md) | クラス | 2026-06-30 14:44 JST |
| [`clsControlSpec.cls`](common/clscontrolspec.cls.md) | クラス | 2026-06-30 14:44 JST |
| [`clsFieldMigrator.cls`](common/clsfieldmigrator.cls.md) | クラス | 2026-06-30 14:44 JST |
| [`clsFieldSpec.cls`](common/clsfieldspec.cls.md) | クラス | 2026-06-30 14:44 JST |
| [`clsFormSpec.cls`](common/clsformspec.cls.md) | クラス | 2026-06-30 14:44 JST |
| [`clsFormatManager.cls`](common/clsformatmanager.cls.md) | クラス | 2026-06-30 14:44 JST |
| [`clsGridIO.cls`](common/clsgridio.cls.md) | クラス | 2026-06-30 14:44 JST |
| [`clsKnowledgeManager.cls`](common/clsknowledgemanager.cls.md) | クラス | 2026-06-30 14:44 JST |
| [`clsLogEntry.cls`](common/clslogentry.cls.md) | クラス | 2026-06-30 14:44 JST |
| [`clsLogger.cls`](common/clslogger.cls.md) | クラス | 2026-06-30 14:44 JST |
| [`clsScreenSpec.cls`](common/clsscreenspec.cls.md) | クラス | 2026-06-30 14:44 JST |
| [`clsSearchEngine.cls`](common/clssearchengine.cls.md) | クラス | 2026-06-30 14:44 JST |
| [`clsSectionSpec.cls`](common/clssectionspec.cls.md) | クラス | 2026-06-30 14:44 JST |
| [`clsSetupOrchestrator.cls`](common/clssetuporchestrator.cls.md) | クラス | 2026-06-30 14:44 JST |
| [`clsSheetRenderer.cls`](common/clssheetrenderer.cls.md) | クラス | 2026-06-30 14:44 JST |
| [`clsStorageResolver.cls`](common/clsstorageresolver.cls.md) | クラス | 2026-06-30 14:44 JST |
| [`clsUserFormRenderer.cls`](common/clsuserformrenderer.cls.md) | クラス | 2026-06-30 14:44 JST |
| [`modBtnGuard.bas`](common/modbtnguard.bas.md) | 標準 | 2026-06-30 14:44 JST |
| [`modBtnMessages.bas`](common/modbtnmessages.bas.md) | 標準 | 2026-06-30 14:44 JST |
| [`modButtonWiring.bas`](common/modbuttonwiring.bas.md) | 標準 | 2026-06-30 14:44 JST |
| [`modCommon.bas`](common/modcommon.bas.md) | 標準 | 2026-06-30 14:44 JST |
| [`modConfigHolder.bas`](common/modconfigholder.bas.md) | 標準 | 2026-06-30 14:44 JST |
| [`modConfigLoader.bas`](common/modconfigloader.bas.md) | 標準 | 2026-06-30 14:44 JST |
| [`modDateUtil.bas`](common/moddateutil.bas.md) | 標準 | 2026-06-30 14:44 JST |
| [`modEntryUserForm.bas`](common/modentryuserform.bas.md) | 標準 | 2026-06-30 14:44 JST |
| [`modFactory.bas`](common/modfactory.bas.md) | 標準 | 2026-06-30 14:44 JST |
| [`modFileIO.bas`](common/modfileio.bas.md) | 標準 | 2026-06-30 14:44 JST |
| [`modFormBuilder.bas`](common/modformbuilder.bas.md) | 標準 | 2026-06-30 14:44 JST |
| [`modFormControlWiring.bas`](common/modformcontrolwiring.bas.md) | 標準 | 2026-06-30 14:44 JST |
| [`modFormatColumns.bas`](common/modformatcolumns.bas.md) | 標準 | 2026-06-30 14:44 JST |
| [`modFormatLoader.bas`](common/modformatloader.bas.md) | 標準 | 2026-06-30 14:44 JST |
| [`modKnowledgeFileIO.bas`](common/modknowledgefileio.bas.md) | 標準 | 2026-06-30 14:44 JST |
| [`modLogIds.bas`](common/modlogids.bas.md) | 標準 | 2026-06-30 14:44 JST |
| [`modPreviewForm.bas`](common/modpreviewform.bas.md) | 標準 | 2026-06-30 14:44 JST |
| [`modRefresh.bas`](common/modrefresh.bas.md) | 標準 | 2026-06-30 15:25 JST |
| [`modScreenRender.bas`](common/modscreenrender.bas.md) | 標準 | 2026-06-30 14:44 JST |
| [`modSetup.bas`](common/modsetup.bas.md) | 標準 | 2026-06-30 14:44 JST |
| [`modSheetButtons.bas`](common/modsheetbuttons.bas.md) | 標準 | 2026-06-30 14:44 JST |
| [`modSheetMap.bas`](common/modsheetmap.bas.md) | 標準 | 2026-06-30 14:44 JST |
| [`modStanzaIO.bas`](common/modstanzaio.bas.md) | 標準 | 2026-06-30 14:44 JST |
| [`modStringUtil.bas`](common/modstringutil.bas.md) | 標準 | 2026-06-30 14:44 JST |
| [`modUILoader.bas`](common/moduiloader.bas.md) | 標準 | 2026-06-30 14:44 JST |
| [`modUserFormCallback.bas`](common/moduserformcallback.bas.md) | 標準 | 2026-06-30 14:44 JST |

