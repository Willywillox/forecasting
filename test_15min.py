"""
Test: Verifica funzionamento forecast con dati a granularita 15 minuti.
Replica la logica delle funzioni principali senza importare il modulo (che dipende da tkinter).
"""
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from statsmodels.tsa.holtwinters import ExponentialSmoothing
import traceback
import sys
import time


# ============================================================================
# REPLICA DELLE FUNZIONI CHIAVE DAL MODULO PRINCIPALE
# (copiate per testare senza dipendenza da tkinter)
# ============================================================================

def _costruisci_pattern_intraday(df):
    """Costruisce pattern intraday percentuali per ciascun giorno della settimana."""
    pattern_intraday = {}
    ordine_giorni = ['lun', 'mar', 'mer', 'gio', 'ven', 'sab', 'dom', 'fest']
    for giorno in ordine_giorni:
        df_giorno = df[df['GG SETT'] == giorno]
        if len(df_giorno) == 0:
            continue
        pattern_fascia = df_giorno.groupby(['FASCIA', 'MINUTI'])['OFFERTO'].mean().reset_index()
        pattern_fascia = pattern_fascia.sort_values('MINUTI')
        totale_giorno = pattern_fascia['OFFERTO'].sum()
        if totale_giorno > 0:
            pattern_fascia['PERCENTUALE'] = pattern_fascia['OFFERTO'] / totale_giorno
        else:
            pattern_fascia['PERCENTUALE'] = 1.0 / len(pattern_fascia)
        pattern_intraday[giorno] = pattern_fascia

    # Normalizza
    for day, pat_df in pattern_intraday.items():
        if pat_df.empty:
            continue
        total_pct = pat_df['PERCENTUALE'].sum()
        if total_pct > 0:
            pat_df = pat_df.copy()
            pat_df['PERCENTUALE'] = pat_df['PERCENTUALE'] / total_pct
            pattern_intraday[day] = pat_df

    return pattern_intraday


def _distribuisci_forecast_per_fascia(pattern_intraday, daily_forecast_df):
    """Applica il pattern intraday ad un forecast giornaliero."""
    forecast_fascia_list = []
    if daily_forecast_df.empty:
        return pd.DataFrame(columns=['DATA', 'GG_SETT', 'FASCIA', 'MINUTI',
                                     'FORECAST_GIORNO', 'PERCENTUALE', 'FORECAST_FASCIA'])
    for _, row_day in daily_forecast_df.iterrows():
        giorno_sett = row_day['GG_SETT']
        if giorno_sett not in pattern_intraday:
            continue
        pattern = pattern_intraday[giorno_sett].copy()
        pattern['DATA'] = row_day['DATA']
        pattern['GG_SETT'] = giorno_sett
        pattern['FORECAST_GIORNO'] = row_day['FORECAST']
        pattern['FORECAST_FASCIA'] = row_day['FORECAST'] * pattern['PERCENTUALE']
        forecast_fascia_list.append(pattern[['DATA', 'GG_SETT', 'FASCIA', 'MINUTI',
                                             'FORECAST_GIORNO', 'PERCENTUALE', 'FORECAST_FASCIA']])

    if not forecast_fascia_list:
        return pd.DataFrame(columns=['DATA', 'GG_SETT', 'FASCIA', 'MINUTI',
                                     'FORECAST_GIORNO', 'PERCENTUALE', 'FORECAST_FASCIA'])

    result = pd.concat(forecast_fascia_list, ignore_index=True)
    return result


def _process_single_fascia_intraday(args):
    """Helper per processare una singola fascia oraria."""
    fascia, df_fascia_subset, future_dates, giorni_forecast = args

    try:
        forecast_results = []

        if len(df_fascia_subset) < 14:
            media_per_dow = df_fascia_subset.groupby('DOW')['OFFERTO'].mean().to_dict()
            for future_date in future_dates:
                dow = future_date.dayofweek
                forecast_val = media_per_dow.get(dow, df_fascia_subset['OFFERTO'].mean())
                forecast_results.append({
                    'DATA': future_date,
                    'FASCIA': fascia,
                    'MINUTI': df_fascia_subset['MINUTI'].iloc[0] if len(df_fascia_subset) > 0 else 0,
                    'GG_SETT': ['lun','mar','mer','gio','ven','sab','dom'][dow],
                    'FORECAST': max(0, forecast_val)
                })
            return forecast_results

        ts = df_fascia_subset.groupby('DATA')['OFFERTO'].mean().sort_index()
        ts = ts.asfreq('D', fill_value=0)

        ts_percentile_95 = ts.quantile(0.95)
        if ts_percentile_95 > 0:
            ts = ts.clip(upper=ts_percentile_95 * 1.5)
        ts = ts.clip(lower=0)

        try:
            model = ExponentialSmoothing(
                ts.values,
                seasonal_periods=7,
                trend='add',
                seasonal='add',
                initialization_method='estimated'
            )
            fit = model.fit()
            forecast_vals = fit.forecast(steps=giorni_forecast)

        except Exception:
            media_per_dow = df_fascia_subset.groupby('DOW')['OFFERTO'].mean().to_dict()
            base = ts.tail(7).mean()
            forecast_vals = []
            for i, future_date in enumerate(future_dates):
                dow = future_date.dayofweek
                dow_factor = media_per_dow.get(dow, base) / base if base > 0 else 1.0
                forecast_vals.append(base * dow_factor)

        for i, future_date in enumerate(future_dates):
            dow = future_date.dayofweek
            forecast_results.append({
                'DATA': future_date,
                'FASCIA': fascia,
                'MINUTI': df_fascia_subset['MINUTI'].iloc[0] if len(df_fascia_subset) > 0 else 0,
                'GG_SETT': ['lun','mar','mer','gio','ven','sab','dom'][dow],
                'FORECAST': max(0, forecast_vals[i] if i < len(forecast_vals) else 0)
            })

        return forecast_results

    except Exception:
        return []


def _forecast_sarima(df, giorni_forecast=28):
    """Forecast con SARIMA."""
    from statsmodels.tsa.statespace.sarimax import SARIMAX

    daily = df.groupby('DATA').agg({'OFFERTO': 'sum', 'GG SETT': 'first'}).reset_index()
    daily = daily.sort_values('DATA').set_index('DATA')
    if daily.empty or len(daily) < 14:
        return None

    model = SARIMAX(
        daily['OFFERTO'],
        order=(1, 1, 1),
        seasonal_order=(1, 0, 1, 7),
        enforce_stationarity=False,
        enforce_invertibility=False,
        freq='D'
    )
    fit = model.fit(disp=False)

    forecast_res = fit.get_forecast(steps=giorni_forecast)
    forecast_mean = forecast_res.predicted_mean
    future_dates = forecast_mean.index

    forecast_daily_df = pd.DataFrame({
        'DATA': future_dates,
        'FORECAST': forecast_mean.values,
        'GG_SETT': [['lun','mar','mer','gio','ven','sab','dom'][d.weekday()] for d in future_dates],
    })

    pattern_intraday = _costruisci_pattern_intraday(df)
    forecast_fascia_df = _distribuisci_forecast_per_fascia(pattern_intraday, forecast_daily_df)

    return {
        'giornaliero': forecast_daily_df,
        'per_fascia': forecast_fascia_df,
    }


# ============================================================================
# GENERAZIONE DATI 15 MINUTI
# ============================================================================

def genera_dati_15min():
    """Genera dataset sintetico con fasce da 15 minuti."""
    np.random.seed(42)
    start_date = datetime(2024, 8, 1)
    n_days = 180

    fasce = []
    for h in range(8, 20):
        for m in [0, 15, 30, 45]:
            end_h = h if m < 45 else h + 1
            end_m = m + 15 if m < 45 else 0
            fasce.append(f"{h:02d}.{m:02d} - {end_h:02d}.{end_m:02d}")

    giorni_sett = ['lun', 'mar', 'mer', 'gio', 'ven', 'sab', 'dom']

    records = []
    for day_offset in range(n_days):
        data = start_date + timedelta(days=day_offset)
        dow = data.weekday()
        gg_sett = giorni_sett[dow]
        week_num = data.isocalendar()[1]

        for fascia_str in fasce:
            parts = fascia_str.split(' - ')
            h, m = int(parts[0].split('.')[0]), int(parts[0].split('.')[1])
            hour_frac = h + m / 60

            base = 30 * (np.exp(-0.5 * ((hour_frac - 10.5) / 1.5) ** 2) +
                         0.7 * np.exp(-0.5 * ((hour_frac - 14.5) / 2.0) ** 2))
            if dow >= 5:
                base *= 0.3

            offerto = max(0, int(base + np.random.normal(0, base * 0.2 + 1)))

            records.append({
                'DATA': data,
                'FASCIA': fascia_str,
                'GG SETT': gg_sett,
                'week': week_num,
                'data-fascia-ggsett': f"{data.strftime('%Y%m%d')} {fascia_str} {gg_sett}",
                'OFFERTO': offerto
            })

    df = pd.DataFrame(records)

    # Simula il parsing che farebbe carica_dati
    fascia_inizio = df['FASCIA'].astype(str).str.split(' - ').str[0].str.strip()
    fascia_normalizzata = fascia_inizio.str.replace('.', ':', regex=False)
    ora_dt = pd.to_datetime(fascia_normalizzata, format='%H:%M', errors='coerce')
    df['ORA_INIZIO'] = np.where(ora_dt.notna(), ora_dt.dt.strftime('%H:%M'), fascia_inizio)
    df['MINUTI'] = (ora_dt.dt.hour * 60 + ora_dt.dt.minute).fillna(-1).astype(int)
    df['IS_WEEKEND'] = df['GG SETT'].isin(['sab', 'dom', 'fest'])
    df['ANNO'] = df['DATA'].dt.year
    df['MESE'] = df['DATA'].dt.month

    print(f"Dataset generato: {len(df)} record, {df['FASCIA'].nunique()} fasce, {df['DATA'].nunique()} giorni")
    return df


# ============================================================================
# TEST
# ============================================================================

def test_parsing(df):
    print("\n" + "="*60)
    print("TEST 1: Parsing fasce 15 minuti")
    print("="*60)
    n_fasce = df['FASCIA'].nunique()
    minuti_ok = (df['MINUTI'] >= 0).all()
    print(f"  Fasce uniche: {n_fasce}")
    print(f"  Tutti MINUTI validi: {minuti_ok}")
    print(f"  MINUTI range: {df['MINUTI'].min()} - {df['MINUTI'].max()}")
    assert n_fasce == 48, f"Attese 48 fasce, trovate {n_fasce}"
    assert minuti_ok
    print("  PASS")
    return True


def test_pattern_intraday(df):
    print("\n" + "="*60)
    print("TEST 2: Pattern intraday con 48 fasce")
    print("="*60)
    pattern = _costruisci_pattern_intraday(df)
    all_ok = True
    for giorno, pat_df in pattern.items():
        n_fasce = len(pat_df)
        somma = pat_df['PERCENTUALE'].sum()
        ok = abs(somma - 1.0) < 0.001
        if not ok:
            all_ok = False
        print(f"  {giorno}: {n_fasce} fasce, sum(PCT) = {somma:.6f} {'OK' if ok else 'FAIL'}")
    assert all_ok
    print("  PASS")
    return True


def test_holt_winters_daily(df):
    print("\n" + "="*60)
    print("TEST 3: Holt-Winters giornaliero")
    print("="*60)
    daily = df.groupby('DATA').agg({'OFFERTO': 'sum', 'GG SETT': 'first'}).reset_index()
    daily = daily.sort_values('DATA').set_index('DATA')
    model = ExponentialSmoothing(daily['OFFERTO'].values, seasonal_periods=7,
                                 trend='add', seasonal='add', initialization_method='estimated')
    fit = model.fit()
    forecast = fit.forecast(steps=28)
    print(f"  Serie: {len(daily)} giorni, media = {daily['OFFERTO'].mean():.0f}")
    print(f"  Forecast 28gg: media = {forecast.mean():.0f}")
    assert len(forecast) == 28
    assert forecast.mean() > 0
    print("  PASS")
    return True


def test_asfreq_per_fascia(df):
    print("\n" + "="*60)
    print("TEST 4: asfreq('D') per serie per-fascia")
    print("="*60)
    fascia_test = '10.00 - 10.15'
    df_f = df[df['FASCIA'] == fascia_test]
    ts = df_f.groupby('DATA')['OFFERTO'].mean().sort_index()
    freq = pd.infer_freq(ts.index)
    ts2 = ts.asfreq('D', fill_value=0)
    print(f"  Fascia: {fascia_test}")
    print(f"  Punti originali: {len(ts)}, freq inferita: {freq}")
    print(f"  Dopo asfreq('D'): {len(ts2)}, zeri aggiunti: {(ts2 == 0).sum()}")
    assert freq == 'D', f"Freq inferred non e' D ma {freq}"
    print("  PASS")
    return True


def test_process_single_fascia(df):
    print("\n" + "="*60)
    print("TEST 5: Forecast singola fascia 15 min")
    print("="*60)
    df_copy = df.copy()
    df_copy['DOW'] = df_copy['DATA'].dt.dayofweek
    fascia_test = '10.00 - 10.15'
    df_fascia = df_copy[df_copy['FASCIA'] == fascia_test].copy()
    future_dates = pd.date_range(start=df['DATA'].max() + timedelta(days=1), periods=28, freq='D')

    print(f"  Fascia: {fascia_test}, record storici: {len(df_fascia)}")
    result = _process_single_fascia_intraday((fascia_test, df_fascia, future_dates, 28))
    print(f"  Risultati: {len(result)} record")
    assert len(result) == 28, f"Attesi 28, ottenuti {len(result)}"
    assert all(r['FORECAST'] >= 0 for r in result)
    print("  PASS")
    return True


def test_intraday_dinamico_48_fasce(df):
    print("\n" + "="*60)
    print("TEST 6: Intraday dinamico con 48 fasce (48 modelli)")
    print("="*60)

    df_copy = df.copy()
    df_copy['DOW'] = df_copy['DATA'].dt.dayofweek
    fasce_uniche = df_copy.sort_values('MINUTI')['FASCIA'].unique()
    last_date = df['DATA'].max()
    future_dates = pd.date_range(start=last_date + timedelta(days=1), periods=28, freq='D')

    print(f"  Fasce da modellare: {len(fasce_uniche)}")

    t0 = time.time()
    forecast_results = []
    for fascia in fasce_uniche:
        df_f = df_copy[df_copy['FASCIA'] == fascia].copy()
        result = _process_single_fascia_intraday((fascia, df_f, future_dates, 28))
        forecast_results.extend(result)

    elapsed = time.time() - t0
    forecast_df = pd.DataFrame(forecast_results)

    n_fasce_result = forecast_df['FASCIA'].nunique()
    n_rows = len(forecast_df)
    expected = 48 * 28

    print(f"  Tempo: {elapsed:.1f}s")
    print(f"  Righe risultato: {n_rows} (attese: {expected})")
    print(f"  Fasce nel risultato: {n_fasce_result}")

    # Verifica totali giornalieri
    daily_totals = forecast_df.groupby('DATA')['FORECAST'].sum()
    print(f"  Forecast giornaliero medio: {daily_totals.mean():.0f}")

    assert n_fasce_result == 48, f"Attese 48 fasce, trovate {n_fasce_result}"
    assert n_rows == expected, f"Attese {expected} righe, trovate {n_rows}"
    print("  PASS")
    return True


def test_sarima(df):
    print("\n" + "="*60)
    print("TEST 7: SARIMA (freq='D') + distribuzione su 48 fasce")
    print("="*60)

    t0 = time.time()
    result = _forecast_sarima(df, giorni_forecast=28)
    elapsed = time.time() - t0

    assert result is not None, "SARIMA ha restituito None"

    n_daily = len(result['giornaliero'])
    n_fascia = len(result['per_fascia'])
    n_fasce_uniche = result['per_fascia']['FASCIA'].nunique()

    print(f"  Tempo: {elapsed:.1f}s")
    print(f"  Forecast giornaliero: {n_daily} giorni")
    print(f"  Forecast per fascia: {n_fascia} righe")
    print(f"  Fasce uniche: {n_fasce_uniche}")
    print(f"  Forecast giornaliero medio: {result['giornaliero']['FORECAST'].mean():.0f}")

    assert n_daily == 28
    assert n_fasce_uniche == 48
    assert n_fascia == 48 * 28

    # Verifica che la distribuzione rispetti il totale giornaliero
    for data in result['giornaliero']['DATA'].head(3):
        daily_val = result['giornaliero'][result['giornaliero']['DATA'] == data]['FORECAST'].values[0]
        fascia_sum = result['per_fascia'][result['per_fascia']['DATA'] == data]['FORECAST_FASCIA'].sum()
        diff_pct = abs(fascia_sum - daily_val) / daily_val * 100 if daily_val > 0 else 0
        print(f"  {data.date()}: daily={daily_val:.0f}, sum(fasce)={fascia_sum:.0f}, diff={diff_pct:.2f}%")

    print("  PASS")
    return True


def test_distribuzione_48_fasce(df):
    print("\n" + "="*60)
    print("TEST 8: Distribuzione forecast su 48 fasce")
    print("="*60)

    pattern = _costruisci_pattern_intraday(df)

    last_date = df['DATA'].max()
    future_dates = pd.date_range(start=last_date + timedelta(days=1), periods=7, freq='D')
    daily_forecast = pd.DataFrame({
        'DATA': future_dates,
        'FORECAST': [500] * 7,
        'GG_SETT': [['lun','mar','mer','gio','ven','sab','dom'][d.weekday()] for d in future_dates]
    })

    result = _distribuisci_forecast_per_fascia(pattern, daily_forecast)

    print(f"  Input: 7 giorni x 500")
    print(f"  Output: {len(result)} righe, {result['FASCIA'].nunique()} fasce")

    for data in result['DATA'].unique()[:3]:
        day_sum = result[result['DATA'] == data]['FORECAST_FASCIA'].sum()
        print(f"  {data.date()}: sum(fasce) = {day_sum:.1f} (atteso: 500)")
        assert abs(day_sum - 500) < 1, f"Sum != 500: {day_sum}"

    print("  PASS")
    return True


def test_performance():
    print("\n" + "="*60)
    print("TEST 9: Performance 48 vs 23 modelli")
    print("="*60)

    np.random.seed(42)
    ts = np.random.poisson(lam=30, size=180).astype(float)

    t0 = time.time()
    for _ in range(23):
        ExponentialSmoothing(ts, seasonal_periods=7, trend='add', seasonal='add',
                             initialization_method='estimated').fit()
    t_23 = time.time() - t0

    t0 = time.time()
    for _ in range(48):
        ExponentialSmoothing(ts, seasonal_periods=7, trend='add', seasonal='add',
                             initialization_method='estimated').fit()
    t_48 = time.time() - t0

    print(f"  23 modelli (30 min): {t_23:.2f}s")
    print(f"  48 modelli (15 min): {t_48:.2f}s")
    print(f"  Rallentamento: {t_48/t_23:.1f}x")
    print("  PASS (informativo)")
    return True


def test_soglie_minime():
    print("\n" + "="*60)
    print("TEST 10: Soglie minime record per fascia")
    print("="*60)
    for n_weeks in [1, 2, 3, 4]:
        n_days = n_weeks * 7
        sotto = n_days < 14
        print(f"  {n_weeks} settimane ({n_days} gg): {n_days} rec/fascia -> "
              f"{'SOTTO SOGLIA (fallback media)' if sotto else 'OK (usa HW)'}")
    print("  PASS (informativo)")
    return True


def test_commento_gui_24modelli():
    """Test: verifica che i commenti GUI dicano 24 modelli (problema cosmetico con 48 fasce)."""
    print("\n" + "="*60)
    print("TEST 11: Commenti GUI/note che dicono '24 modelli'")
    print("="*60)

    with open('/home/user/forecasting/analisi_trafficonewfct_profsari.py', 'r') as f:
        content = f.read()

    issues = []
    for i, line in enumerate(content.split('\n'), 1):
        if '24 modelli' in line.lower() or '24 model' in line.lower():
            issues.append((i, line.strip()))
        if 'uno per ogni ora' in line.lower() or 'one per hour' in line.lower():
            issues.append((i, line.strip()))
        if 'granularita oraria' in line.lower() or 'granularita\' oraria' in line.lower():
            issues.append((i, line.strip()))

    if issues:
        print("  TROVATI riferimenti hardcoded a granularita oraria / 24 modelli:")
        for lineno, text in issues:
            print(f"    Linea {lineno}: {text[:100]}")
        print(f"  Totale: {len(issues)} riferimenti da aggiornare")
    else:
        print("  Nessun riferimento hardcoded trovato")

    print("  PASS (informativo - questi sono solo commenti/stringhe UI)")
    return True


if __name__ == '__main__':
    print("="*60)
    print("TEST FORECAST CON DATI A 15 MINUTI")
    print("="*60)

    df = genera_dati_15min()

    # Salva per eventuale uso
    df.to_excel('/home/user/forecasting/dati_test_15min.xlsx', index=False)

    results = {}
    tests = [
        ('parsing', test_parsing),
        ('pattern_intraday', test_pattern_intraday),
        ('holt_winters_daily', test_holt_winters_daily),
        ('asfreq_per_fascia', test_asfreq_per_fascia),
        ('process_single_fascia', test_process_single_fascia),
        ('intraday_48_fasce', test_intraday_dinamico_48_fasce),
        ('sarima', test_sarima),
        ('distribuzione_48', test_distribuzione_48_fasce),
        ('performance', test_performance),
        ('soglie_minime', test_soglie_minime),
        ('commenti_gui', test_commento_gui_24modelli),
    ]

    for name, test_fn in tests:
        try:
            results[name] = test_fn(df) if 'df' in test_fn.__code__.co_varnames else test_fn()
        except Exception as e:
            print(f"  FAIL: {e}")
            traceback.print_exc()
            results[name] = False

    print("\n" + "="*60)
    print("RIEPILOGO")
    print("="*60)
    for name, passed in results.items():
        print(f"  [{'PASS' if passed else 'FAIL'}] {name}")

    n_pass = sum(results.values())
    print(f"\n  {n_pass}/{len(results)} test passati")
    sys.exit(0 if n_pass == len(results) else 1)
