import pandas as pd
import matplotlib
matplotlib.use('Agg')  # Use non-interactive backend
import matplotlib.pyplot as plt
import numpy as np
from scipy import stats
import os
import gc
import warnings
from datetime import datetime
warnings.filterwarnings('ignore')

# Set style for better looking plots
plt.style.use('seaborn-v0_8-whitegrid')
plt.rcParams['figure.figsize'] = (14, 8)
plt.rcParams['font.size'] = 10
plt.rcParams['axes.titlesize'] = 12
plt.rcParams['axes.labelsize'] = 10

def load_grouped_data(filepath):
    """Load all sheets from a grouped Excel file"""
    xl = pd.ExcelFile(filepath)
    sheets = {}
    for sheet_name in xl.sheet_names:
        sheets[sheet_name] = pd.read_excel(filepath, sheet_name=sheet_name)
    return sheets

def get_baseline_data(df):
    """Get baseline data for comparison
    For baseline files: use the last 10 minutes mapped to -20 to -10
    For experiment files used as baseline: use post-treatment data starting from minute 3"""
    if 'Minutes_from_Time0' not in df.columns:
        return df
    # Check if this is a baseline file (has negative minutes in -20 to -10 range)
    # or an experiment file (has positive minutes >= 3)
    if len(df[df['Minutes_from_Time0'].between(-20, -10, inclusive='both')]) > 0:
        # Baseline file: use the -20 to -10 range (last 10 minutes of baseline measurement)
        return df[df['Minutes_from_Time0'].between(-20, -10, inclusive='both')].copy()
    else:
        # Experiment file used as baseline: use post-treatment data starting from minute 3
        return df[df['Minutes_from_Time0'] >= 3].copy()

def get_pre_post_data(df):
    """Split experiment data into pre-treatment and post-treatment
    Pre-treatment ignores first 1.5 minutes
    Post-treatment ignores first 2 minutes - starts from minute 3"""
    if 'Minutes_from_Time0' not in df.columns:
        return None, None
    
    # Pre-treatment: ignore first 1.5 minutes (exclude data from -1.5 to 0)
    pre = df[(df['Minutes_from_Time0'] < -1.5)].copy()
    # Start post-treatment from minute 3 (ignore minutes 1 and 2)
    post = df[df['Minutes_from_Time0'] >= 3].copy()
    
    return pre, post

def calculate_stats(data, column):
    """Calculate statistics for a data column"""
    if data is None or len(data) == 0:
        return None
    
    vals = data[column].dropna() if column in data.columns else pd.Series([])
    
    if len(vals) == 0:
        return None
    
    return {
        'mean': vals.mean(),
        'std': vals.std(),
        'n': len(vals),
        'sem': vals.std() / np.sqrt(len(vals)),
        'values': vals
    }

def main():
    # Directory
    directory = os.getcwd()  # Use current working directory
    
    # Find the most recent grouped files
    import glob
    baseline_files = sorted(glob.glob(os.path.join(directory, "*antidote baseline 301225_grouped.xlsx")), reverse=True)
    experiment_files = sorted(glob.glob(os.path.join(directory, "*antidote 301225_grouped.xlsx")), reverse=True)
    
    if not baseline_files:
        print("ERROR: No baseline grouped file found!")
        return
    if not experiment_files:
        print("ERROR: No experiment grouped file found!")
        return
    
    baseline_file = baseline_files[0]
    experiment_file = experiment_files[0]
    
    # Load data
    print("Loading data...")
    print(f"  Baseline: {os.path.basename(baseline_file)}")
    print(f"  Experiment: {os.path.basename(experiment_file)}")
    baseline_data = load_grouped_data(baseline_file)
    experiment_data = load_grouped_data(experiment_file)
    
    print(f"Baseline groups: {list(baseline_data.keys())}")
    print(f"Experiment groups: {list(experiment_data.keys())}")
    
    # Get common groups
    groups = sorted(set(baseline_data.keys()) & set(experiment_data.keys()))
    print(f"Common groups for comparison: {groups}")
    
    # Key measurement columns to analyze
    key_columns = ['f', 'TVb', 'MVb', 'Penh', 'Ti', 'Te', 'PIFb', 'PEFb']
    
    # Generate timestamp prefix for file names (yyyy-mm-dd hhmmss)
    timestamp = datetime.now().strftime("%Y-%m-%d %H%M%S")
    
    # Create output directory for plots
    plots_dir = os.path.join(directory, "analysis_plots_v3")
    os.makedirs(plots_dir, exist_ok=True)
    
    # ========================================================================
    # Data preparation: Baseline (post only) vs Experiment (pre and post)
    # ========================================================================
    print("\n" + "="*70)
    print("DATA STRUCTURE:")
    print("  - BASELINE: Post-treatment data only, starting from minute 3 (reference)")
    print("  - EXPERIMENT PRE: Before treatment (first 1.5 minutes excluded)")
    print("  - EXPERIMENT POST: After treatment, starting from minute 3")
    print("  NOTE: Pre-treatment ignores first 1.5 minutes (data before -1.5 minutes)")
    print("  NOTE: Minutes 1-2 are excluded from post-treatment analysis")
    print("="*70)
    
    # Store all statistics
    all_stats = {}
    
    for group in groups:
        all_stats[group] = {}
        
        # Get baseline (post-treatment only)
        baseline = get_baseline_data(baseline_data[group])
        
        # Get experiment pre and post
        exp_pre, exp_post = get_pre_post_data(experiment_data[group])
        
        for col in key_columns:
            bl_stats = calculate_stats(baseline, col)
            pre_stats = calculate_stats(exp_pre, col)
            post_stats = calculate_stats(exp_post, col)
            
            all_stats[group][col] = {
                'baseline': bl_stats,
                'exp_pre': pre_stats,
                'exp_post': post_stats
            }
    
    # ========================================================================
    # PART 1: Summary Statistics
    # ========================================================================
    print("\n" + "="*70)
    print("PART 1: SUMMARY STATISTICS")
    print("="*70)
    
    for group in groups:
        print(f"\n{'='*50}")
        print(f"  {group.upper()}")
        print(f"{'='*50}")
        print(f"{'Parameter':<10} {'Baseline':>12} {'Exp Pre':>12} {'Exp Post':>12} {'Pre vs BL':>12} {'Post vs BL':>12}")
        print("-" * 70)
        
        for col in key_columns:
            col_stats = all_stats[group][col]
            
            bl_mean = col_stats['baseline']['mean'] if col_stats['baseline'] else float('nan')
            pre_mean = col_stats['exp_pre']['mean'] if col_stats['exp_pre'] else float('nan')
            post_mean = col_stats['exp_post']['mean'] if col_stats['exp_post'] else float('nan')
            
            # Calculate percent difference from baseline
            pre_vs_bl = ((pre_mean - bl_mean) / bl_mean * 100) if bl_mean and not np.isnan(bl_mean) else float('nan')
            post_vs_bl = ((post_mean - bl_mean) / bl_mean * 100) if bl_mean and not np.isnan(bl_mean) else float('nan')
            
            print(f"{col:<10} {bl_mean:>12.2f} {pre_mean:>12.2f} {post_mean:>12.2f} {pre_vs_bl:>+11.1f}% {post_vs_bl:>+11.1f}%")
    
    # ========================================================================
    # PART 2: Statistical Comparisons
    # ========================================================================
    print("\n" + "="*70)
    print("PART 2: STATISTICAL COMPARISONS")
    print("="*70)
    
    comparison_results = {}
    
    for group in groups:
        print(f"\n{'='*50}")
        print(f"  {group.upper()}")
        print(f"{'='*50}")
        comparison_results[group] = {}
        
        print("\n--- Experiment Pre vs Baseline ---")
        for col in key_columns:
            col_stats = all_stats[group][col]
            if col_stats['baseline'] and col_stats['exp_pre']:
                t_stat, p_val = stats.ttest_ind(col_stats['baseline']['values'], col_stats['exp_pre']['values'])
                diff_pct = ((col_stats['exp_pre']['mean'] - col_stats['baseline']['mean']) / col_stats['baseline']['mean'] * 100)
                sig = "***" if p_val < 0.001 else "**" if p_val < 0.01 else "*" if p_val < 0.05 else ""
                print(f"  {col}: Baseline={col_stats['baseline']['mean']:.2f}, Exp Pre={col_stats['exp_pre']['mean']:.2f}, "
                      f"Diff={diff_pct:+.1f}%, p={p_val:.4f} {sig}")
                comparison_results[group][f'{col}_pre_vs_bl'] = {'p': p_val, 'diff': diff_pct, 'sig': sig}
        
        print("\n--- Experiment Post vs Baseline ---")
        for col in key_columns:
            col_stats = all_stats[group][col]
            if col_stats['baseline'] and col_stats['exp_post']:
                t_stat, p_val = stats.ttest_ind(col_stats['baseline']['values'], col_stats['exp_post']['values'])
                diff_pct = ((col_stats['exp_post']['mean'] - col_stats['baseline']['mean']) / col_stats['baseline']['mean'] * 100)
                sig = "***" if p_val < 0.001 else "**" if p_val < 0.01 else "*" if p_val < 0.05 else ""
                print(f"  {col}: Baseline={col_stats['baseline']['mean']:.2f}, Exp Post={col_stats['exp_post']['mean']:.2f}, "
                      f"Diff={diff_pct:+.1f}%, p={p_val:.4f} {sig}")
                comparison_results[group][f'{col}_post_vs_bl'] = {'p': p_val, 'diff': diff_pct, 'sig': sig}
        
        print("\n--- Experiment Pre vs Post (Treatment Effect) ---")
        for col in key_columns:
            col_stats = all_stats[group][col]
            if col_stats['exp_pre'] and col_stats['exp_post']:
                t_stat, p_val = stats.ttest_ind(col_stats['exp_pre']['values'], col_stats['exp_post']['values'])
                diff_pct = ((col_stats['exp_post']['mean'] - col_stats['exp_pre']['mean']) / col_stats['exp_pre']['mean'] * 100)
                sig = "***" if p_val < 0.001 else "**" if p_val < 0.01 else "*" if p_val < 0.05 else ""
                print(f"  {col}: Pre={col_stats['exp_pre']['mean']:.2f} -> Post={col_stats['exp_post']['mean']:.2f}, "
                      f"Change={diff_pct:+.1f}%, p={p_val:.4f} {sig}")
                comparison_results[group][f'{col}_treatment'] = {'p': p_val, 'diff': diff_pct, 'sig': sig}
    
    # ========================================================================
    # Save Statistical Analysis Summary
    # ========================================================================
    stats_summary_filename = f'{timestamp} statistical_analysis_summary.txt'
    stats_summary_path = os.path.join(directory, stats_summary_filename)
    with open(stats_summary_path, 'w', encoding='utf-8') as f:
        f.write("STATISTICAL ANALYSIS SUMMARY\n")
        f.write("="*70 + "\n")
        f.write(f"Analysis Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write("\n")
        f.write("METHODOLOGY:\n")
        f.write("- Pre-treatment data: First 1.5 minutes excluded (data before -1.5 minutes)\n")
        f.write("- Post-treatment data: First 2 minutes excluded (starts from minute 3)\n")
        f.write("- Statistical test: Independent t-test (two-sample t-test)\n")
        f.write("- Significance levels: * p<0.05, ** p<0.01, *** p<0.001\n")
        f.write("\n" + "="*70 + "\n\n")
        
        for group in groups:
            f.write(f"\n{'='*70}\n")
            f.write(f"GROUP: {group.upper()}\n")
            f.write(f"{'='*70}\n\n")
            
            # Experiment Pre vs Baseline
            f.write("EXPERIMENT PRE-TREATMENT vs BASELINE\n")
            f.write("-"*70 + "\n")
            f.write(f"{'Parameter':<12} {'Baseline Mean':>15} {'Exp Pre Mean':>15} {'Difference %':>15} {'p-value':>12} {'Significance':>12}\n")
            f.write("-"*70 + "\n")
            for col in key_columns:
                key = f'{col}_pre_vs_bl'
                if key in comparison_results[group]:
                    result = comparison_results[group][key]
                    col_stats = all_stats[group][col]
                    bl_mean = col_stats['baseline']['mean'] if col_stats['baseline'] else 0
                    pre_mean = col_stats['exp_pre']['mean'] if col_stats['exp_pre'] else 0
                    f.write(f"{col:<12} {bl_mean:>15.2f} {pre_mean:>15.2f} {result['diff']:>+14.1f}% {result['p']:>12.4f} {result['sig']:>12}\n")
            f.write("\n")
            
            # Experiment Post vs Baseline
            f.write("EXPERIMENT POST-TREATMENT vs BASELINE\n")
            f.write("-"*70 + "\n")
            f.write(f"{'Parameter':<12} {'Baseline Mean':>15} {'Exp Post Mean':>15} {'Difference %':>15} {'p-value':>12} {'Significance':>12}\n")
            f.write("-"*70 + "\n")
            for col in key_columns:
                key = f'{col}_post_vs_bl'
                if key in comparison_results[group]:
                    result = comparison_results[group][key]
                    col_stats = all_stats[group][col]
                    bl_mean = col_stats['baseline']['mean'] if col_stats['baseline'] else 0
                    post_mean = col_stats['exp_post']['mean'] if col_stats['exp_post'] else 0
                    f.write(f"{col:<12} {bl_mean:>15.2f} {post_mean:>15.2f} {result['diff']:>+14.1f}% {result['p']:>12.4f} {result['sig']:>12}\n")
            f.write("\n")
            
            # Treatment Effect (Pre vs Post)
            f.write("TREATMENT EFFECT (PRE vs POST)\n")
            f.write("-"*70 + "\n")
            f.write(f"{'Parameter':<12} {'Pre Mean':>15} {'Post Mean':>15} {'Change %':>15} {'p-value':>12} {'Significance':>12}\n")
            f.write("-"*70 + "\n")
            for col in key_columns:
                key = f'{col}_treatment'
                if key in comparison_results[group]:
                    result = comparison_results[group][key]
                    col_stats = all_stats[group][col]
                    pre_mean = col_stats['exp_pre']['mean'] if col_stats['exp_pre'] else 0
                    post_mean = col_stats['exp_post']['mean'] if col_stats['exp_post'] else 0
                    f.write(f"{col:<12} {pre_mean:>15.2f} {post_mean:>15.2f} {result['diff']:>+14.1f}% {result['p']:>12.4f} {result['sig']:>12}\n")
            f.write("\n")
    
    print(f"\nStatistical analysis summary saved to: {stats_summary_path}")
    
    # ========================================================================
    # VISUALIZATION 1: Bar Chart - Baseline vs Exp Pre vs Exp Post
    # ========================================================================
    print("\n\nCreating visualizations...")
    
    for group in groups:
        fig, axes = plt.subplots(2, 4, figsize=(14, 8))
        axes = axes.flatten()
        
        for idx, col in enumerate(key_columns):
            ax = axes[idx]
            col_stats = all_stats[group][col]
            
            x = np.arange(3)
            
            means = [
                col_stats['baseline']['mean'] if col_stats['baseline'] else 0,
                col_stats['exp_pre']['mean'] if col_stats['exp_pre'] else 0,
                col_stats['exp_post']['mean'] if col_stats['exp_post'] else 0
            ]
            sems = [
                col_stats['baseline']['sem'] if col_stats['baseline'] else 0,
                col_stats['exp_pre']['sem'] if col_stats['exp_pre'] else 0,
                col_stats['exp_post']['sem'] if col_stats['exp_post'] else 0
            ]
            
            colors = ['steelblue', 'lightsalmon', 'coral']
            bars = ax.bar(x, means, yerr=sems, capsize=4, color=colors, edgecolor='black', linewidth=1.2)
            
            ax.set_ylabel(col)
            ax.set_title(f'{col}')
            ax.set_xticks(x)
            ax.set_xticklabels(['Baseline', 'Exp Pre', 'Exp Post'], rotation=15)
            
            # Add significance markers
            if col_stats['exp_pre'] and col_stats['baseline']:
                _, p = stats.ttest_ind(col_stats['baseline']['values'], col_stats['exp_pre']['values'])
                if p < 0.05:
                    max_y = max(means[0], means[1]) + max(sems[0], sems[1])
                    stars = "***" if p < 0.001 else "**" if p < 0.01 else "*"
                    ax.plot([0, 1], [max_y * 1.05, max_y * 1.05], 'k-', linewidth=1)
                    ax.text(0.5, max_y * 1.07, stars, ha='center', fontsize=10)
            
            if col_stats['exp_post'] and col_stats['baseline']:
                _, p = stats.ttest_ind(col_stats['baseline']['values'], col_stats['exp_post']['values'])
                if p < 0.05:
                    max_y = max(means[0], means[2]) + max(sems[0], sems[2])
                    stars = "***" if p < 0.001 else "**" if p < 0.01 else "*"
                    ax.plot([0, 2], [max_y * 1.15, max_y * 1.15], 'k-', linewidth=1)
                    ax.text(1, max_y * 1.17, stars, ha='center', fontsize=10)
        
        plt.suptitle(f'{group.upper()}: Baseline vs Experiment (Pre/Post Treatment)', fontsize=14, fontweight='bold')
        plt.tight_layout()
        filename = f'{timestamp} comparison_{group.replace(" ", "_")}.png'
        plt.savefig(os.path.join(plots_dir, filename), dpi=100, bbox_inches='tight')
        plt.close()
        gc.collect()
        print(f"  Saved: {filename}")
    
    # ========================================================================
    # VISUALIZATION 2: Time Course with Baseline Reference (All Groups Combined)
    # ========================================================================
    print("\nCreating time-course plots...")
    
    # Load processed data (individual subjects) for statistical testing
    print("  Loading processed data for statistical analysis...")
    baseline_proc_files = sorted(glob.glob(os.path.join(directory, "*antidote baseline 301225_processed.xlsx")), reverse=True)
    exp_proc_files = sorted(glob.glob(os.path.join(directory, "*antidote 301225_processed.xlsx")), reverse=True)
    
    baseline_processed = load_grouped_data(baseline_proc_files[0]) if baseline_proc_files else {}
    exp_processed = load_grouped_data(exp_proc_files[0]) if exp_proc_files else {}
    
    # Map sheet names to groups (e.g., '1a.WBPth' -> 'group a')
    def get_group_letter(sheet_name):
        import re
        match = re.search(r'\d+([a-zA-Z])', sheet_name)
        if match:
            return f"group {match.group(1).lower()}"
        return None
    
    # Use all key columns for timecourse plots
    selected_cols = ['f', 'TVb', 'MVb', 'Penh', 'Ti', 'Te', 'PIFb', 'PEFb']
    # Colors for all possible groups
    group_colors = {
        'group a': '#E74C3C',  # Red
        'group b': '#3498DB',  # Blue
        'group c': '#27AE60',  # Green
        'group d': '#9B59B6',  # Purple
        'group e': '#F39C12',  # Orange
        'group f': '#1ABC9C',  # Teal
        'group g': '#E91E63',  # Pink
        'group h': '#00BCD4',  # Cyan
    }
    group_markers = {
        'group a': 'o', 
        'group b': 's', 
        'group c': '^',
        'group d': 'D',  # Diamond
        'group e': 'v',  # Triangle down
        'group f': 'p',  # Pentagon
        'group g': 'h',  # Hexagon
        'group h': '*',  # Star
    }
    
    # Split into two figures: first 4 and last 4 parameters
    param_sets = [
        (selected_cols[:4], 'timecourse_vs_baseline_1.png', 'Time Course (Part 1): f, TVb, MVb, Penh'),
        (selected_cols[4:], 'timecourse_vs_baseline_2.png', 'Time Course (Part 2): Ti, Te, PIFb, PEFb'),
    ]
    
    for param_set, filename_suffix, figure_title in param_sets:
        fig, axes = plt.subplots(2, 2, figsize=(16, 12))
        axes = axes.flatten()
        
        for j, col in enumerate(param_set):
            ax = axes[j]
            
            # Store data for statistical testing
            timepoint_data = {}  # {minute: {group: [values]}}
            
            for group in groups:
                color = group_colors.get(group, 'gray')
                marker = group_markers.get(group, 'o')
                baseline_full = baseline_data[group]
                experiment = experiment_data[group]
                
                all_times = []
                all_means = []
                all_sems = []
                
                # Get baseline data in the -20 to -10 range (last 10 minutes of baseline measurement)
                if 'Minutes_from_Time0' in baseline_full.columns and col in baseline_full.columns:
                    # Check if this is a baseline file (has negative minutes) or experiment file
                    baseline_range = baseline_full[baseline_full['Minutes_from_Time0'].between(-20, -10, inclusive='both')].copy()
                    if len(baseline_range) == 0:
                        # Fallback: if no -20 to -10 range, try to get last 10 minutes from post-treatment data
                        baseline_post = baseline_full[baseline_full['Minutes_from_Time0'] >= 3].copy()
                        if len(baseline_post) > 0:
                            max_time = baseline_post['Minutes_from_Time0'].max()
                            baseline_range = baseline_post[baseline_post['Minutes_from_Time0'] >= max_time - 10].copy()
                            if len(baseline_range) > 0:
                                # Map to -20 to -10 range
                                orig_min = baseline_range['Minutes_from_Time0'].min()
                                orig_max = baseline_range['Minutes_from_Time0'].max()
                                if orig_max > orig_min:
                                    baseline_range['mapped_time'] = -20 + (baseline_range['Minutes_from_Time0'] - orig_min) / (orig_max - orig_min) * 10
                                else:
                                    baseline_range['mapped_time'] = -15
                                baseline_range['Minutes_from_Time0'] = baseline_range['mapped_time']
                    
                    if len(baseline_range) > 0:
                        # Group by minute and calculate mean/SEM
                        baseline_range['minute'] = baseline_range['Minutes_from_Time0'].round().astype(int)
                        for minute, grp in baseline_range.groupby('minute'):
                            vals = grp[col].dropna()
                            if len(vals) > 0:
                                all_times.append(minute)
                                all_means.append(vals.mean())
                                all_sems.append(vals.std() / np.sqrt(len(vals)) if len(vals) > 1 else 0)
                                # Store for stats
                                if minute not in timepoint_data:
                                    timepoint_data[minute] = {}
                                timepoint_data[minute][group] = vals.values
                
                # Get experiment data (limit to 30 min post-treatment)
                if col in experiment.columns and 'Minutes_from_Time0' in experiment.columns:
                    exp_data = experiment[(experiment['Minutes_from_Time0'] <= 30)].copy()
                    exp_data['minute'] = exp_data['Minutes_from_Time0'].round().astype(int)
                    
                    for minute, grp in exp_data.groupby('minute'):
                        vals = grp[col].dropna()
                        if len(vals) > 0:
                            all_times.append(minute)
                            all_means.append(vals.mean())
                            all_sems.append(vals.std() / np.sqrt(len(vals)) if len(vals) > 1 else 0)
                            # Store for stats
                            if minute not in timepoint_data:
                                timepoint_data[minute] = {}
                            timepoint_data[minute][group] = vals.values
                
                # Sort by time
                if all_times:
                    sorted_idx = np.argsort(all_times)
                    all_times = np.array(all_times)[sorted_idx]
                    all_means = np.array(all_means)[sorted_idx]
                    all_sems = np.array(all_sems)[sorted_idx]
                    
                    # Plot with error bars
                    ax.errorbar(all_times, all_means, yerr=all_sems, 
                               fmt=f'-{marker}', color=color, 
                               label=group.replace('group ', 'Group ').upper(),
                               markersize=5, alpha=0.8, linewidth=1.5, capsize=2, capthick=1)
            
            # Collect data from individual subjects (processed files) for statistical testing
            # This gives us multiple observations per group per timepoint
            stats_data = {}  # {minute: {group: [values from all subjects]}}
            
            # Collect experiment data from processed file
            for sheet_name, sheet_df in exp_processed.items():
                grp = get_group_letter(sheet_name)
                if grp is None or col not in sheet_df.columns or 'Minutes_from_Time0' not in sheet_df.columns:
                    continue
                
                # Get data limited to 30 min post-treatment
                exp_sheet = sheet_df[(sheet_df['Minutes_from_Time0'] <= 30)].copy()
                exp_sheet['minute'] = exp_sheet['Minutes_from_Time0'].round().astype(int)
                
                for minute, grp_rows in exp_sheet.groupby('minute'):
                    vals = grp_rows[col].dropna().values
                    if len(vals) > 0:
                        if minute not in stats_data:
                            stats_data[minute] = {}
                        if grp not in stats_data[minute]:
                            stats_data[minute][grp] = []
                        stats_data[minute][grp].extend(vals)
            
            # Perform statistical tests at each timepoint (Kruskal-Wallis - non-parametric)
            significant_points = []
            test_count = 0
            sig_count = 0
            for minute in sorted(stats_data.keys()):
                grp_data = stats_data[minute]
                # Need at least 2 groups with data to compare
                valid_groups = [g for g in groups if g in grp_data and len(grp_data[g]) >= 2]
                if len(valid_groups) >= 2:
                    try:
                        # Use Kruskal-Wallis H-test (non-parametric alternative to one-way ANOVA)
                        data_arrays = [np.array(grp_data[g]) for g in valid_groups]
                        test_count += 1
                        stat, p_val = stats.kruskal(*data_arrays)
                        if p_val < 0.05:
                            sig_count += 1
                            # Get y position for marker (max of all group means at this timepoint)
                            y_vals = [np.mean(grp_data[g]) for g in valid_groups]
                            significant_points.append((minute, max(y_vals), p_val))
                    except Exception as e:
                        pass
            
            print(f"    {col}: Tested {test_count} timepoints, found {sig_count} significant")
            if sig_count > 0:
                for minute, y_pos, p_val in significant_points:
                    stars = '***' if p_val < 0.001 else '**' if p_val < 0.01 else '*'
                    print(f"      -> Minute {minute}: p={p_val:.4f} {stars}")
            
            # Mark significant timepoints with visible markers
            for minute, y_pos, p_val in significant_points:
                stars = '***' if p_val < 0.001 else '**' if p_val < 0.01 else '*'
                # Use black asterisks with no background
                ax.annotate(stars, xy=(minute, y_pos), xytext=(minute, y_pos * 1.08),
                           ha='center', fontsize=10, color='black', fontweight='bold')
            
            # Add vertical lines
            ax.axvline(x=0, color='black', linestyle='--', linewidth=2, label='Treatment')
            ax.axvline(x=-10, color='navy', linestyle='--', linewidth=1.5, alpha=0.7)
            
            # Set x-axis limits: -25 to 32
            ax.set_xlim(-25, 32)
            
            # Add shading: blue (-20 to -10), green (-10 to 0), red (0 to 30)
            ax.axvspan(-20, -10, alpha=0.2, color='blue', label='Baseline Period')
            ax.axvspan(-10, 0, alpha=0.2, color='green')
            ax.axvspan(0, 30, alpha=0.2, color='red')
            
            ax.set_xlabel('Minutes from Treatment', fontsize=10)
            ax.set_ylabel(col, fontsize=10)
            ax.set_title(f'{col} - All Groups Time Course', fontsize=12, fontweight='bold')
            
            if j == 0:
                ax.legend(loc='upper right', fontsize=8)
        
        plt.suptitle(f'{figure_title}\n(* p<0.05, ** p<0.01, *** p<0.001 - Kruskal-Wallis test between groups)', 
                     fontsize=12, fontweight='bold')
        plt.tight_layout()
        filename = f'{timestamp} {filename_suffix}'
        plt.savefig(os.path.join(plots_dir, filename), dpi=150, bbox_inches='tight')
        plt.close()
        gc.collect()
        print(f"  Saved: {filename}")
    
    # ========================================================================
    # VISUALIZATION 3: Difference from Baseline Heatmap
    # ========================================================================
    print("\nCreating difference heatmap...")
    
    fig, axes = plt.subplots(1, 2, figsize=(12, 4))
    
    # Pre vs Baseline
    pre_diffs = []
    for group in groups:
        row = []
        for col in key_columns:
            col_stats = all_stats[group][col]
            if col_stats['baseline'] and col_stats['exp_pre']:
                diff = ((col_stats['exp_pre']['mean'] - col_stats['baseline']['mean']) / col_stats['baseline']['mean'] * 100)
            else:
                diff = 0
            row.append(diff)
        pre_diffs.append(row)
    
    im1 = axes[0].imshow(pre_diffs, cmap='RdBu_r', aspect='auto', vmin=-200, vmax=200)
    axes[0].set_xticks(np.arange(len(key_columns)))
    axes[0].set_yticks(np.arange(len(groups)))
    axes[0].set_xticklabels(key_columns, rotation=45, ha='right')
    axes[0].set_yticklabels([g.replace('group ', '').upper() for g in groups])
    axes[0].set_title('EXPERIMENT PRE vs BASELINE (%)', fontweight='bold')
    for i in range(len(groups)):
        for j in range(len(key_columns)):
            val = pre_diffs[i][j]
            color = 'white' if abs(val) > 100 else 'black'
            axes[0].text(j, i, f'{val:+.0f}%', ha='center', va='center', color=color, fontsize=9)
    plt.colorbar(im1, ax=axes[0], label='% Difference from Baseline')
    
    # Post vs Baseline
    post_diffs = []
    for group in groups:
        row = []
        for col in key_columns:
            col_stats = all_stats[group][col]
            if col_stats['baseline'] and col_stats['exp_post']:
                diff = ((col_stats['exp_post']['mean'] - col_stats['baseline']['mean']) / col_stats['baseline']['mean'] * 100)
            else:
                diff = 0
            row.append(diff)
        post_diffs.append(row)
    
    im2 = axes[1].imshow(post_diffs, cmap='RdBu_r', aspect='auto', vmin=-200, vmax=200)
    axes[1].set_xticks(np.arange(len(key_columns)))
    axes[1].set_yticks(np.arange(len(groups)))
    axes[1].set_xticklabels(key_columns, rotation=45, ha='right')
    axes[1].set_yticklabels([g.replace('group ', '').upper() for g in groups])
    axes[1].set_title('EXPERIMENT POST vs BASELINE (%)', fontweight='bold')
    for i in range(len(groups)):
        for j in range(len(key_columns)):
            val = post_diffs[i][j]
            color = 'white' if abs(val) > 100 else 'black'
            axes[1].text(j, i, f'{val:+.0f}%', ha='center', va='center', color=color, fontsize=9)
    plt.colorbar(im2, ax=axes[1], label='% Difference from Baseline')
    
    plt.tight_layout()
    filename = f'{timestamp} diff_from_baseline_heatmap.png'
    plt.savefig(os.path.join(plots_dir, filename), dpi=100, bbox_inches='tight')
    plt.close()
    gc.collect()
    print(f"  Saved: {filename}")
    
    # ========================================================================
    # VISUALIZATION 4: Treatment Effect Lines
    # ========================================================================
    print("\nCreating treatment effect plots...")
    
    # Use all key columns for treatment effect plots
    selected_cols_treatment = ['f', 'TVb', 'MVb', 'Penh', 'Ti', 'Te', 'PIFb', 'PEFb']
    
    fig, axes = plt.subplots(len(groups), 8, figsize=(20, 3*len(groups)))
    if len(groups) == 1:
        axes = axes.reshape(1, -1)
    
    for i, group in enumerate(groups):
        for j, col in enumerate(selected_cols_treatment):
            ax = axes[i, j]
            col_stats = all_stats[group][col]
            
            if col_stats['baseline'] and col_stats['exp_pre'] and col_stats['exp_post']:
                bl_mean = col_stats['baseline']['mean']
                pre_mean = col_stats['exp_pre']['mean']
                post_mean = col_stats['exp_post']['mean']
                
                # Plot baseline as reference line
                ax.axhline(y=bl_mean, color='steelblue', linestyle='-', linewidth=3, label='Baseline')
                
                # Plot experiment pre to post
                ax.plot([0, 1], [pre_mean, post_mean], 'o-', color='coral', 
                       linewidth=2.5, markersize=12, label='Experiment')
                
                # Add percentage annotations
                pre_vs_bl = ((pre_mean - bl_mean) / bl_mean * 100)
                post_vs_bl = ((post_mean - bl_mean) / bl_mean * 100)
                treatment_effect = ((post_mean - pre_mean) / pre_mean * 100) if pre_mean != 0 else 0
                
                ax.annotate(f'{pre_vs_bl:+.0f}% vs BL', xy=(0, pre_mean), xytext=(-0.15, pre_mean),
                           fontsize=9, color='coral', fontweight='bold')
                ax.annotate(f'{post_vs_bl:+.0f}% vs BL', xy=(1, post_mean), xytext=(1.05, post_mean),
                           fontsize=9, color='coral', fontweight='bold')
                ax.annotate(f'Treatment:\n{treatment_effect:+.0f}%', xy=(0.5, (pre_mean+post_mean)/2),
                           fontsize=8, ha='center', color='darkred')
            
            ax.set_xticks([0, 1])
            ax.set_xticklabels(['Pre-Treatment', 'Post-Treatment'])
            ax.set_ylabel(col)
            ax.set_title(f'{group.upper()}: {col}')
            ax.set_xlim(-0.3, 1.5)
            
            if i == 0 and j == 0:
                ax.legend(loc='best', fontsize=8)
    
    plt.tight_layout()
    filename = f'{timestamp} treatment_effect_lines.png'
    plt.savefig(os.path.join(plots_dir, filename), dpi=100, bbox_inches='tight')
    plt.close()
    gc.collect()
    print(f"  Saved: {filename}")
    
    # ========================================================================
    # VISUALIZATION 5: Normalized to Baseline (100%)
    # ========================================================================
    print("\nCreating normalized comparison...")
    
    fig, axes = plt.subplots(2, 4, figsize=(14, 8))
    axes = axes.flatten()
    
    for idx, col in enumerate(key_columns):
        ax = axes[idx]
        
        x = np.arange(len(groups))
        width = 0.25
        
        # Baseline is always 100%
        baseline_normalized = [100] * len(groups)
        
        pre_normalized = []
        post_normalized = []
        
        for group in groups:
            col_stats = all_stats[group][col]
            if col_stats['baseline'] and col_stats['exp_pre']:
                pre_normalized.append((col_stats['exp_pre']['mean'] / col_stats['baseline']['mean']) * 100)
            else:
                pre_normalized.append(0)
            
            if col_stats['baseline'] and col_stats['exp_post']:
                post_normalized.append((col_stats['exp_post']['mean'] / col_stats['baseline']['mean']) * 100)
            else:
                post_normalized.append(0)
        
        ax.bar(x - width, baseline_normalized, width, label='Baseline', color='steelblue', edgecolor='black')
        ax.bar(x, pre_normalized, width, label='Exp Pre', color='lightsalmon', edgecolor='black')
        ax.bar(x + width, post_normalized, width, label='Exp Post', color='coral', edgecolor='black')
        
        ax.axhline(y=100, color='gray', linestyle='--', alpha=0.5)
        
        ax.set_ylabel(f'{col} (% of Baseline)')
        ax.set_title(f'{col}')
        ax.set_xticks(x)
        ax.set_xticklabels([g.replace('group ', '').upper() for g in groups])
        
        if idx == 0:
            ax.legend(loc='best', fontsize=8)
    
    plt.suptitle('All Parameters Normalized to Baseline (100%)', fontsize=14, fontweight='bold')
    plt.tight_layout()
    filename = f'{timestamp} normalized_to_baseline.png'
    plt.savefig(os.path.join(plots_dir, filename), dpi=100, bbox_inches='tight')
    plt.close()
    gc.collect()
    print(f"  Saved: {filename}")
    
    # ========================================================================
    # VISUALIZATION 6: Summary Grouped Bar Chart
    # ========================================================================
    print("\nCreating summary chart...")
    
    fig, axes = plt.subplots(2, 4, figsize=(14, 8))
    axes = axes.flatten()
    
    for idx, col in enumerate(key_columns):
        ax = axes[idx]
        
        x = np.arange(len(groups))
        width = 0.25
        
        baseline_vals = []
        pre_vals = []
        post_vals = []
        baseline_sems = []
        pre_sems = []
        post_sems = []
        
        for group in groups:
            col_stats = all_stats[group][col]
            baseline_vals.append(col_stats['baseline']['mean'] if col_stats['baseline'] else 0)
            pre_vals.append(col_stats['exp_pre']['mean'] if col_stats['exp_pre'] else 0)
            post_vals.append(col_stats['exp_post']['mean'] if col_stats['exp_post'] else 0)
            baseline_sems.append(col_stats['baseline']['sem'] if col_stats['baseline'] else 0)
            pre_sems.append(col_stats['exp_pre']['sem'] if col_stats['exp_pre'] else 0)
            post_sems.append(col_stats['exp_post']['sem'] if col_stats['exp_post'] else 0)
        
        ax.bar(x - width, baseline_vals, width, yerr=baseline_sems, label='Baseline', 
               color='steelblue', edgecolor='black', capsize=3)
        ax.bar(x, pre_vals, width, yerr=pre_sems, label='Exp Pre', 
               color='lightsalmon', edgecolor='black', capsize=3)
        ax.bar(x + width, post_vals, width, yerr=post_sems, label='Exp Post', 
               color='coral', edgecolor='black', capsize=3)
        
        ax.set_ylabel(col)
        ax.set_title(f'{col} Comparison')
        ax.set_xticks(x)
        ax.set_xticklabels([g.replace('group ', '').upper() for g in groups])
        
        if idx == 0:
            ax.legend(loc='best', fontsize=8)
    
    plt.suptitle('Baseline vs Experiment Pre/Post Comparison', fontsize=14, fontweight='bold')
    plt.tight_layout()
    filename = f'{timestamp} summary_comparison.png'
    plt.savefig(os.path.join(plots_dir, filename), dpi=100, bbox_inches='tight')
    plt.close()
    gc.collect()
    print(f"  Saved: {filename}")
    
    # ========================================================================
    # SUMMARY AND CONCLUSIONS
    # ========================================================================
    print("\n" + "="*70)
    print("SUMMARY AND KEY FINDINGS")
    print("="*70)
    
    print("\n[1] EXPERIMENT PRE-TREATMENT vs BASELINE")
    print("-" * 50)
    print("(Shows initial state of experiment animals compared to baseline)")
    
    for group in groups:
        print(f"\n{group.upper()}:")
        for col in key_columns:
            col_stats = all_stats[group][col]
            if col_stats['baseline'] and col_stats['exp_pre']:
                diff = ((col_stats['exp_pre']['mean'] - col_stats['baseline']['mean']) / col_stats['baseline']['mean'] * 100)
                t, p = stats.ttest_ind(col_stats['baseline']['values'], col_stats['exp_pre']['values'])
                sig = "***" if p < 0.001 else "**" if p < 0.01 else "*" if p < 0.05 else ""
                if abs(diff) > 20 or p < 0.05:
                    print(f"  {col}: {diff:+.1f}% vs Baseline (p={p:.4f}) {sig}")
    
    print("\n\n[2] EXPERIMENT POST-TREATMENT vs BASELINE")
    print("-" * 50)
    print("(Shows state after treatment compared to baseline)")
    
    for group in groups:
        print(f"\n{group.upper()}:")
        for col in key_columns:
            col_stats = all_stats[group][col]
            if col_stats['baseline'] and col_stats['exp_post']:
                diff = ((col_stats['exp_post']['mean'] - col_stats['baseline']['mean']) / col_stats['baseline']['mean'] * 100)
                t, p = stats.ttest_ind(col_stats['baseline']['values'], col_stats['exp_post']['values'])
                sig = "***" if p < 0.001 else "**" if p < 0.01 else "*" if p < 0.05 else ""
                if abs(diff) > 20 or p < 0.05:
                    print(f"  {col}: {diff:+.1f}% vs Baseline (p={p:.4f}) {sig}")
    
    print("\n\n[3] TREATMENT EFFECT (Pre to Post)")
    print("-" * 50)
    print("(Shows the effect of treatment within experiment group)")
    
    for group in groups:
        print(f"\n{group.upper()}:")
        for col in key_columns:
            col_stats = all_stats[group][col]
            if col_stats['exp_pre'] and col_stats['exp_post']:
                diff = ((col_stats['exp_post']['mean'] - col_stats['exp_pre']['mean']) / col_stats['exp_pre']['mean'] * 100)
                t, p = stats.ttest_ind(col_stats['exp_pre']['values'], col_stats['exp_post']['values'])
                sig = "***" if p < 0.001 else "**" if p < 0.01 else "*" if p < 0.05 else ""
                if abs(diff) > 20 or p < 0.05:
                    direction = "increased" if diff > 0 else "decreased"
                    print(f"  {col}: {direction} by {abs(diff):.1f}% (p={p:.4f}) {sig}")
    
    print("\n\n[4] KEY CONCLUSIONS")
    print("-" * 50)
    
    # Analyze key parameters
    print("\nPenh (Airway Resistance):")
    for group in groups:
        penh_stats = all_stats[group]['Penh']
        if penh_stats['baseline'] and penh_stats['exp_pre'] and penh_stats['exp_post']:
            pre_vs_bl = ((penh_stats['exp_pre']['mean'] - penh_stats['baseline']['mean']) / penh_stats['baseline']['mean'] * 100)
            post_vs_bl = ((penh_stats['exp_post']['mean'] - penh_stats['baseline']['mean']) / penh_stats['baseline']['mean'] * 100)
            print(f"  {group.upper()}: Pre={pre_vs_bl:+.0f}% vs BL, Post={post_vs_bl:+.0f}% vs BL")
            if pre_vs_bl > 100 and post_vs_bl < pre_vs_bl:
                print(f"    -> Treatment REDUCED Penh towards baseline")
    
    print("\nMVb (Minute Ventilation):")
    for group in groups:
        mvb_stats = all_stats[group]['MVb']
        if mvb_stats['baseline'] and mvb_stats['exp_pre'] and mvb_stats['exp_post']:
            pre_vs_bl = ((mvb_stats['exp_pre']['mean'] - mvb_stats['baseline']['mean']) / mvb_stats['baseline']['mean'] * 100)
            post_vs_bl = ((mvb_stats['exp_post']['mean'] - mvb_stats['baseline']['mean']) / mvb_stats['baseline']['mean'] * 100)
            print(f"  {group.upper()}: Pre={pre_vs_bl:+.0f}% vs BL, Post={post_vs_bl:+.0f}% vs BL")
            if pre_vs_bl < -30 and post_vs_bl > pre_vs_bl:
                print(f"    -> Treatment INCREASED MVb towards baseline")
    
    print("\n" + "="*70)
    print(f"All visualizations saved to: {plots_dir}")
    print("="*70)
    
    # Save summary to file
    summary_filename = f'{timestamp} analysis_summary_v3.txt'
    summary_path = os.path.join(directory, summary_filename)
    with open(summary_path, 'w', encoding='utf-8') as f:
        f.write("EXPERIMENT ANALYSIS SUMMARY (v3)\n")
        f.write("================================\n")
        f.write("Analysis compares: Baseline | Experiment Pre | Experiment Post\n")
        f.write("Baseline = post-treatment data from baseline condition (reference)\n")
        f.write("NOTE: Pre-treatment analysis ignores first 1.5 minutes (data before -1.5 minutes)\n")
        f.write("NOTE: Post-treatment analysis starts from minute 3 (minutes 1-2 excluded)\n")
        f.write("NOTE: Timeframe graphs show up to 30 minutes post-treatment\n\n")
        
        f.write("="*70 + "\n")
        f.write("SUMMARY STATISTICS\n")
        f.write("="*70 + "\n\n")
        
        for group in groups:
            f.write(f"\n{group.upper()}\n")
            f.write("-"*50 + "\n")
            f.write(f"{'Parameter':<10} {'Baseline':>12} {'Exp Pre':>12} {'Exp Post':>12} {'Pre vs BL':>12} {'Post vs BL':>12}\n")
            
            for col in key_columns:
                col_stats = all_stats[group][col]
                bl_mean = col_stats['baseline']['mean'] if col_stats['baseline'] else 0
                pre_mean = col_stats['exp_pre']['mean'] if col_stats['exp_pre'] else 0
                post_mean = col_stats['exp_post']['mean'] if col_stats['exp_post'] else 0
                pre_vs_bl = ((pre_mean - bl_mean) / bl_mean * 100) if bl_mean else 0
                post_vs_bl = ((post_mean - bl_mean) / bl_mean * 100) if bl_mean else 0
                f.write(f"{col:<10} {bl_mean:>12.2f} {pre_mean:>12.2f} {post_mean:>12.2f} {pre_vs_bl:>+11.1f}% {post_vs_bl:>+11.1f}%\n")
    
    print(f"\nSummary saved to: {summary_path}")

if __name__ == "__main__":
    main()

