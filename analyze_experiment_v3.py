import pandas as pd
import matplotlib
matplotlib.use('Agg')  # Use non-interactive backend
import matplotlib.pyplot as plt
import numpy as np
from scipy import stats
import os
import gc
import warnings
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
    """Get baseline data (post-treatment only, as reference)"""
    if 'Minutes_from_Time0' not in df.columns:
        return df
    # Use post-treatment data as baseline reference
    return df[df['Minutes_from_Time0'] > 0].copy()

def get_pre_post_data(df):
    """Split experiment data into pre-treatment and post-treatment"""
    if 'Minutes_from_Time0' not in df.columns:
        return None, None
    
    pre = df[df['Minutes_from_Time0'] < 0].copy()
    post = df[df['Minutes_from_Time0'] > 0].copy()
    
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
    directory = r"C:\Users\user\ambs"
    
    # Load data
    print("Loading data...")
    baseline_data = load_grouped_data(os.path.join(directory, "baseline_grouped.xlsx"))
    experiment_data = load_grouped_data(os.path.join(directory, "exp_grouped.xlsx"))
    
    print(f"Baseline groups: {list(baseline_data.keys())}")
    print(f"Experiment groups: {list(experiment_data.keys())}")
    
    # Get common groups
    groups = sorted(set(baseline_data.keys()) & set(experiment_data.keys()))
    print(f"Common groups for comparison: {groups}")
    
    # Key measurement columns to analyze
    key_columns = ['f', 'TVb', 'MVb', 'Penh', 'Ti', 'Te', 'PIFb', 'PEFb']
    
    # Create output directory for plots
    plots_dir = os.path.join(directory, "analysis_plots_v3")
    os.makedirs(plots_dir, exist_ok=True)
    
    # ========================================================================
    # Data preparation: Baseline (post only) vs Experiment (pre and post)
    # ========================================================================
    print("\n" + "="*70)
    print("DATA STRUCTURE:")
    print("  - BASELINE: Post-treatment data only (used as reference)")
    print("  - EXPERIMENT PRE: Before treatment")
    print("  - EXPERIMENT POST: After treatment")
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
        plt.savefig(os.path.join(plots_dir, f'comparison_{group.replace(" ", "_")}.png'), dpi=100, bbox_inches='tight')
        plt.close()
        gc.collect()
        print(f"  Saved: comparison_{group.replace(' ', '_')}.png")
    
    # ========================================================================
    # VISUALIZATION 2: Time Course with Baseline Reference
    # ========================================================================
    print("\nCreating time-course plots...")
    
    fig, axes = plt.subplots(len(groups), 4, figsize=(14, 3*len(groups)))
    if len(groups) == 1:
        axes = axes.reshape(1, -1)
    
    selected_cols = ['f', 'TVb', 'Penh', 'MVb']
    
    for i, group in enumerate(groups):
        baseline = get_baseline_data(baseline_data[group])
        experiment = experiment_data[group]
        
        for j, col in enumerate(selected_cols):
            ax = axes[i, j]
            
            # Plot experiment time course
            if col in experiment.columns and 'Minutes_from_Time0' in experiment.columns:
                time_vals = experiment['Minutes_from_Time0'].values
                col_vals = experiment[col].values
                ax.plot(time_vals, col_vals, '-o', color='coral', label='Experiment', 
                       markersize=4, alpha=0.8, linewidth=1.5)
            
            # Plot baseline as horizontal band (mean +/- SEM)
            if col in baseline.columns:
                bl_mean = baseline[col].mean()
                bl_sem = baseline[col].std() / np.sqrt(len(baseline))
                ax.axhline(y=bl_mean, color='steelblue', linestyle='-', linewidth=2, label='Baseline')
                ax.axhspan(bl_mean - bl_sem, bl_mean + bl_sem, alpha=0.3, color='steelblue')
            
            # Mark treatment time
            ax.axvline(x=0, color='black', linestyle='--', linewidth=2, label='Treatment')
            
            # Add shading
            ax.axvspan(ax.get_xlim()[0] if ax.get_xlim()[0] < 0 else -10, 0, alpha=0.1, color='green')
            ax.axvspan(0, ax.get_xlim()[1] if ax.get_xlim()[1] > 0 else 30, alpha=0.1, color='red')
            
            ax.set_xlabel('Minutes from Treatment')
            ax.set_ylabel(col)
            ax.set_title(f'{group.upper()}: {col}')
            
            if i == 0 and j == 0:
                ax.legend(loc='best', fontsize=7)
    
    plt.tight_layout()
    plt.savefig(os.path.join(plots_dir, 'timecourse_vs_baseline.png'), dpi=100, bbox_inches='tight')
    plt.close()
    gc.collect()
    print(f"  Saved: timecourse_vs_baseline.png")
    
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
    plt.savefig(os.path.join(plots_dir, 'diff_from_baseline_heatmap.png'), dpi=100, bbox_inches='tight')
    plt.close()
    gc.collect()
    print(f"  Saved: diff_from_baseline_heatmap.png")
    
    # ========================================================================
    # VISUALIZATION 4: Treatment Effect Lines
    # ========================================================================
    print("\nCreating treatment effect plots...")
    
    fig, axes = plt.subplots(len(groups), 4, figsize=(14, 3*len(groups)))
    if len(groups) == 1:
        axes = axes.reshape(1, -1)
    
    selected_cols = ['f', 'TVb', 'Penh', 'MVb']
    
    for i, group in enumerate(groups):
        for j, col in enumerate(selected_cols):
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
    plt.savefig(os.path.join(plots_dir, 'treatment_effect_lines.png'), dpi=100, bbox_inches='tight')
    plt.close()
    gc.collect()
    print(f"  Saved: treatment_effect_lines.png")
    
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
    plt.savefig(os.path.join(plots_dir, 'normalized_to_baseline.png'), dpi=100, bbox_inches='tight')
    plt.close()
    gc.collect()
    print(f"  Saved: normalized_to_baseline.png")
    
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
    plt.savefig(os.path.join(plots_dir, 'summary_comparison.png'), dpi=100, bbox_inches='tight')
    plt.close()
    gc.collect()
    print(f"  Saved: summary_comparison.png")
    
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
    summary_path = os.path.join(directory, "analysis_summary_v3.txt")
    with open(summary_path, 'w', encoding='utf-8') as f:
        f.write("EXPERIMENT ANALYSIS SUMMARY (v3)\n")
        f.write("================================\n")
        f.write("Analysis compares: Baseline | Experiment Pre | Experiment Post\n")
        f.write("Baseline = post-treatment data from baseline condition (reference)\n\n")
        
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

