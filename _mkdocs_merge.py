# -*- coding: utf-8 -*-
"""
Merge local 7757b37 modules nav into remote mkdocs.yml,
remove now-deleted file references (setup.md, operations.md, customize.md, source/*).

Strategy:
  - Start with remote mkdocs.yml
  - Replace its nav: section with a clean new nav that references only files that exist
  - Keep header (site_name through extra_css) as remote
"""
import io
import os
import subprocess
import sys

PUSH_DIR = r"C:\kvba\push"
BK_DIR = os.path.join(PUSH_DIR, "_bk")
LOG_PATH = os.path.join(PUSH_DIR, "_mkdocs_merge.log")

logf = open(LOG_PATH, "w", encoding="utf-8")
def L(msg):
    logf.write(str(msg) + "\n")
    logf.flush()

L("=== _mkdocs_merge.py START ===")

# Step 1: get remote mkdocs.yml content (bytes from git show)
os.chdir(PUSH_DIR)
remote_yml = subprocess.check_output(["git", "show", "origin/main:mkdocs.yml"])
remote_text = remote_yml.decode("utf-8")
L(f"remote yml byte len={len(remote_yml)}, text len={len(remote_text)}")

# Step 2: load local mkdocs.yml
with open(os.path.join(PUSH_DIR, "mkdocs.yml"), "rb") as f:
    local_yml = f.read()
# detect BOM
if local_yml.startswith(b'\xef\xbb\xbf'):
    local_text = local_yml[3:].decode("utf-8")
else:
    local_text = local_yml.decode("utf-8")
L(f"local yml byte len={len(local_yml)}, text len={len(local_text)}")

# Step 3: detect nav start in remote
remote_lines = remote_text.split("\n")
nav_start = -1
for i, line in enumerate(remote_lines):
    if line.startswith("nav:"):
        nav_start = i
        break
L(f"remote nav: starts at line {nav_start}")

# header = everything before nav
header = "\n".join(remote_lines[:nav_start])
L(f"header lines={nav_start}")

# Step 4: build new nav
# Files that exist in our planned final tree:
# - From remote (we keep): index.md, architecture.md, spec.md, settings.md,
#   screen_customization_guide_v2.md, source/**, stylesheets/extra.css
#   (we keep setup.md, operations.md - WAIT - they're deleted in remote!)
# Re-check: confirmed remote_files - what's deleted vs kept
# Files NOT in remote: setup.md, operations.md, customize.md (and others)
# Files we are adding: install.md, usage.md, troubleshooting.md, modules/**

# We need a nav that references only files that will exist after our commit.
# Build the new nav text manually.

new_nav = """nav:
  - 概要: index.md
  - インストール: install.md
  - 使い方: usage.md
  - 困ったとき: troubleshooting.md
  - 設計:
      - 全体アーキテクチャ: architecture.md
      - 仕様: spec.md
  - 設定とカスタマイズ:
      - 設定値の変更: settings.md
      - 画面のカスタマイズ: screen_customization_guide_v2.md
  - モジュール一覧:
      - 一覧: modules/index.md
      - 管理.xlsm 用:
          - ThisWorkbook.cls: modules/admin/thisworkbook.cls.md
          - clsFormatDesignScreen.cls: modules/admin/clsformatdesignscreen.cls.md
          - clsFormatListScreen.cls: modules/admin/clsformatlistscreen.cls.md
          - clsFormatPreviewScreen.cls: modules/admin/clsformatpreviewscreen.cls.md
          - clsLogScreen.cls: modules/admin/clslogscreen.cls.md
          - clsMigrationScreen.cls: modules/admin/clsmigrationscreen.cls.md
          - clsStorageScreen.cls: modules/admin/clsstoragescreen.cls.md
          - clsSystemSettingsScreen.cls: modules/admin/clssystemsettingsscreen.cls.md
          - modEntryFormat.bas: modules/admin/modentryformat.bas.md
          - modEntrySettings.bas: modules/admin/modentrysettings.bas.md
      - 登録修正.xlsm 用:
          - ThisWorkbook.cls: modules/register/thisworkbook.cls.md
          - clsKnowledgeEditScreen.cls: modules/register/clsknowledgeeditscreen.cls.md
          - clsKnowledgeRegisterScreen.cls: modules/register/clsknowledgeregisterscreen.cls.md
          - modEntryKnowledge.bas: modules/register/modentryknowledge.bas.md
          - modKnowledgeEntryHelpers.bas: modules/register/modknowledgeentryhelpers.bas.md
      - 検索.xlsm 用:
          - ThisWorkbook.cls: modules/search/thisworkbook.cls.md
          - clsKnowledgeViewScreen.cls: modules/search/clsknowledgeviewscreen.cls.md
          - clsSearchScreen.cls: modules/search/clssearchscreen.cls.md
          - modEntrySearch.bas: modules/search/modentrysearch.bas.md
      - 共通モジュール:
          - clsButtonSpec.cls: modules/common/clsbuttonspec.cls.md
          - clsCellAddrHelper.cls: modules/common/clscelladdrhelper.cls.md
          - clsCellBinding.cls: modules/common/clscellbinding.cls.md
          - clsCellIO.cls: modules/common/clscellio.cls.md
          - clsControlSpec.cls: modules/common/clscontrolspec.cls.md
          - clsFieldMigrator.cls: modules/common/clsfieldmigrator.cls.md
          - clsFieldSpec.cls: modules/common/clsfieldspec.cls.md
          - clsFormatManager.cls: modules/common/clsformatmanager.cls.md
          - clsFormSpec.cls: modules/common/clsformspec.cls.md
          - clsGridIO.cls: modules/common/clsgridio.cls.md
          - clsKnowledgeManager.cls: modules/common/clsknowledgemanager.cls.md
          - clsLogEntry.cls: modules/common/clslogentry.cls.md
          - clsLogger.cls: modules/common/clslogger.cls.md
          - clsScreenSpec.cls: modules/common/clsscreenspec.cls.md
          - clsSearchEngine.cls: modules/common/clssearchengine.cls.md
          - clsSectionSpec.cls: modules/common/clssectionspec.cls.md
          - clsSetupOrchestrator.cls: modules/common/clssetuporchestrator.cls.md
          - clsSheetRenderer.cls: modules/common/clssheetrenderer.cls.md
          - clsStanzaSection.cls: modules/common/clsstanzasection.cls.md
          - clsStanzaValidationIssue.cls: modules/common/clsstanzavalidationissue.cls.md
          - clsStanzaValidationResult.cls: modules/common/clsstanzavalidationresult.cls.md
          - clsStorageResolver.cls: modules/common/clsstorageresolver.cls.md
          - clsUserFormRenderer.cls: modules/common/clsuserformrenderer.cls.md
          - IScreenRenderer.cls: modules/common/iscreenrenderer.cls.md
          - modButtonWiring.bas: modules/common/modbuttonwiring.bas.md
          - modCommon.bas: modules/common/modcommon.bas.md
          - modConfigHolder.bas: modules/common/modconfigholder.bas.md
          - modConfigLoader.bas: modules/common/modconfigloader.bas.md
          - modDateUtil.bas: modules/common/moddateutil.bas.md
          - modE2E_AllButtons.bas: modules/common/mode2e_allbuttons.bas.md
          - modEntryUserForm.bas: modules/common/modentryuserform.bas.md
          - modFactory.bas: modules/common/modfactory.bas.md
          - modFileIO.bas: modules/common/modfileio.bas.md
          - modFormatColumns.bas: modules/common/modformatcolumns.bas.md
          - modFormatLoader.bas: modules/common/modformatloader.bas.md
          - modFormBuilder.bas: modules/common/modformbuilder.bas.md
          - modKnowledgeFileIO.bas: modules/common/modknowledgefileio.bas.md
          - modLogIds.bas: modules/common/modlogids.bas.md
          - modRefresh.bas: modules/common/modrefresh.bas.md
          - modScreenRender.bas: modules/common/modscreenrender.bas.md
          - modSetup.bas: modules/common/modsetup.bas.md
          - modSheetButtons.bas: modules/common/modsheetbuttons.bas.md
          - modSheetMap.bas: modules/common/modsheetmap.bas.md
          - modSpecExamples.bas: modules/common/modspecexamples.bas.md
          - modStanzaIO.bas: modules/common/modstanzaio.bas.md
          - modStringUtil.bas: modules/common/modstringutil.bas.md
          - modTestHelpers.bas: modules/common/modtesthelpers.bas.md
          - modUILoader.bas: modules/common/moduiloader.bas.md
          - modUserFormCallback.bas: modules/common/moduserformcallback.bas.md
  - ソースコード (旧):
      - 一覧: source/index.md
      - インストール層:
          - modSetup.bas: source/infrastructure/modsetup.md
      - エントリポイント層:
          - modEntryFormat.bas: source/entrypoint/modentryformat.md
          - modEntryKnowledge.bas: source/entrypoint/modentryknowledge.md
          - modEntrySearch.bas: source/entrypoint/modentrysearch.md
          - modEntrySettings.bas: source/entrypoint/modentrysettings.md
          - modSpecExamples.bas: source/entrypoint/modspecexamples.md
      - 画面層:
          - clsBackupMgmtScreen.cls: source/screen/clsbackupmgmtscreen.md
          - clsErrorBandScreen.cls: source/screen/clserrorbandscreen.md
          - clsFileFormatScreen.cls: source/screen/clsfileformatscreen.md
          - clsFormatDesignScreen.cls: source/screen/clsformatdesignscreen.md
          - clsFormatListScreen.cls: source/screen/clsformatlistscreen.md
          - clsFormatPreviewScreen.cls: source/screen/clsformatpreviewscreen.md
          - clsKnowledgeEditScreen.cls: source/screen/clsknowledgeeditscreen.md
          - clsKnowledgeListScreen.cls: source/screen/clsknowledgelistscreen.md
          - clsKnowledgeRegisterScreen.cls: source/screen/clsknowledgeregisterscreen.md
          - clsKnowledgeViewScreen.cls: source/screen/clsknowledgeviewscreen.md
          - clsLogScreen.cls: source/screen/clslogscreen.md
          - clsMigrationScreen.cls: source/screen/clsmigrationscreen.md
          - clsSearchScreen.cls: source/screen/clssearchscreen.md
          - clsStorageScreen.cls: source/screen/clsstoragescreen.md
          - clsSystemSettingsScreen.cls: source/screen/clssystemsettingsscreen.md
      - ビジネスロジック層:
          - clsButtonSpec.cls: source/business-logic/clsbuttonspec.md
          - clsControlSpec.cls: source/business-logic/clscontrolspec.md
          - clsFieldMigrator.cls: source/business-logic/clsfieldmigrator.md
          - clsFieldSpec.cls: source/business-logic/clsfieldspec.md
          - clsFormatManager.cls: source/business-logic/clsformatmanager.md
          - clsFormSpec.cls: source/business-logic/clsformspec.md
          - clsKnowledgeManager.cls: source/business-logic/clsknowledgemanager.md
          - clsLogger.cls: source/business-logic/clslogger.md
          - clsScreenSpec.cls: source/business-logic/clsscreenspec.md
          - clsSearchEngine.cls: source/business-logic/clssearchengine.md
          - clsSectionSpec.cls: source/business-logic/clssectionspec.md
          - clsSetupOrchestrator.cls: source/business-logic/clssetuporchestrator.md
          - clsSheetRenderer.cls: source/business-logic/clssheetrenderer.md
          - clsStorageResolver.cls: source/business-logic/clsstorageresolver.md
          - clsUserFormRenderer.cls: source/business-logic/clsuserformrenderer.md
          - IScreenRenderer.cls: source/business-logic/iscreenrenderer.md
          - modFactory.bas: source/business-logic/modfactory.md
          - modFormBuilder.bas: source/business-logic/modformbuilder.md
          - modScreenRender.bas: source/business-logic/modscreenrender.md
      - ユーティリティ層:
          - clsCellAddrHelper.cls: source/utility/clscelladdrhelper.md
          - clsCellBinding.cls: source/utility/clscellbinding.md
          - clsCellIO.cls: source/utility/clscellio.md
          - clsGridIO.cls: source/utility/clsgridio.md
          - clsLogEntry.cls: source/utility/clslogentry.md
          - ClsStanzaSection.cls: source/utility/clsstanzasection.md
          - ClsStanzaValidationIssue.cls: source/utility/clsstanzavalidationissue.md
          - ClsStanzaValidationResult.cls: source/utility/clsstanzavalidationresult.md
          - modCommon.bas: source/utility/modcommon.md
          - modConfigHolder.bas: source/utility/modconfigholder.md
          - modConfigLoader.bas: source/utility/modconfigloader.md
          - modDateUtil.bas: source/utility/moddateutil.md
          - modFileIO.bas: source/utility/modfileio.md
          - modFormatColumns.bas: source/utility/modformatcolumns.md
          - modFormatLoader.bas: source/utility/modformatloader.md
          - modImageRender.bas: source/utility/modimagerender.md
          - modKnowledgeFileIO.bas: source/utility/modknowledgefileio.md
          - modLogIds.bas: source/utility/modlogids.md
          - modSheetMap.bas: source/utility/modsheetmap.md
          - modStanzaIO.bas: source/utility/modstanzaio.md
          - modStringUtil.bas: source/utility/modstringutil.md
          - modUIConfig.bas: source/utility/moduiconfig.md
          - modUILayoutExtractor.bas: source/utility/moduilayoutextractor.md
          - modUILoader.bas: source/utility/moduiloader.md
      - 特殊モジュール:
          - ThisWorkbook（検索.xlsm）: source/special/thisworkbook-search.md
          - ThisWorkbook（登録修正.xlsm）: source/special/thisworkbook-register.md
          - ThisWorkbook（管理.xlsm）: source/special/thisworkbook-admin.md
"""

merged = header + "\n" + new_nav

# Save merged
merged_path = os.path.join(PUSH_DIR, "_bk", "mkdocs_merged.yml")
with open(merged_path, "w", encoding="utf-8", newline="\n") as f:
    f.write(merged)
L(f"merged saved to {merged_path}, len={len(merged)}")

L("=== _mkdocs_merge.py DONE ===")
logf.close()
print("OK")
