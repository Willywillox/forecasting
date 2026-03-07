# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all, collect_submodules

# Collect all data/binaries/hiddenimports for problematic packages
prophet_datas, prophet_bins, prophet_hidden = collect_all('prophet')
tbats_datas, tbats_bins, tbats_hidden = collect_all('tbats')
cmdstanpy_datas, cmdstanpy_bins, cmdstanpy_hidden = collect_all('cmdstanpy')
holidays_datas, holidays_bins, holidays_hidden = collect_all('holidays')
statsmodels_hidden = collect_submodules('statsmodels')
scipy_hidden = collect_submodules('scipy')

all_datas = prophet_datas + tbats_datas + cmdstanpy_datas + holidays_datas
all_binaries = prophet_bins + tbats_bins + cmdstanpy_bins + holidays_bins
all_hidden = (
    prophet_hidden + tbats_hidden + cmdstanpy_hidden + holidays_hidden
    + statsmodels_hidden + scipy_hidden
    + ['sklearn', 'sklearn.linear_model', 'sklearn.preprocessing']
)

a = Analysis(
    ['analisi_trafficonewfct_profsari.py'],
    pathex=[],
    binaries=all_binaries,
    datas=all_datas,
    hiddenimports=all_hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='AnalisiForecast',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='AnalisiForecast',
)
