import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from scipy import stats
import os
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

def get_pre_post_data(df, exclude_time_zero=True):
    """Split data into pre-treatment and post-treatment"""
    if 'Minutes_from_Time0' not in df.columns:
        return None, None
    
    if exclude_time_zero:
        pre = df[df['Minutes_from_Time0'] < 0].copy()
        post = df[df['Minutes_from_Time0'] > 0].copy()
    else:
        pre = df[df['Minutes_from_Time0'] <= 0].copy()
        post = df[df['Minutes_from_Time0'] >= 0].copy()
    
    return pre, post

def calculate_pre_post_stats(df, column):
    """Calculate pre vs post treatment statistics for a single dataframe"""
    pre, post = get_pre_post_data(df)
    
    if pre is None or post is None:
        return None
    
    pre_vals = pre[column].dropna() if column in pre.columns else pd.Series([])
    post_vals = post[column].dropna() if column in post.columns else pd.Series([])
    
    if len(pre_vals) == 0 or len(post_vals) == 0:
        return None
    
    pre_mean = pre_vals.mean()
    post_mean = post_vals.mean()
    pre_std = pre_vals.std()
    post_std = post_vals.std()
    
    # Percent change from pre to post
    if pre_mean != 0:
        pct_change = ((post_mean - pre_mean) / abs(pre_mean)) * 100
    else:
        pct_change = float('inf') if post_mean != 0 else 0
    
    # Paired t-test (using means per time point would be better, but we'll use unpaired here)
    try:
        t_stat, p_value = stats.ttest_ind(pre_vals, post_vals)
    except:
        t_stat, p_value = None, None
    
    return {
        'pre_mean': pre_mean,
        'post_mean': post_mean,
        'pre_std': pre_std,
        'post_std': post_std,
        'pre_n': len(pre_vals),
        'post_n': len(post_vals),
        'pct_change': pct_change,
        't_stat': t_stat,
        'p_value': p_value
    }

def calculate_treatment_effect(baseline_df, exp_df, column):
    """
    Calculate the treatment effect by comparing:
    - Pre-to-post change in baseline
    - Pre-to-post change in experiment
    """
    baseline_stats = calculate_pre_post_stats(baseline_df, column)
    exp_stats = calculate_pre_post_stats(exp_df, column)
    
    if baseline_stats is None or exp_stats is None:
        return None
    
    # Delta-delta analysis: (Exp_post - Exp_pre) - (Baseline_post - Baseline_pre)
    baseline_delta = baseline_stats['post_mean'] - baseline_stats['pre_mean']
    exp_delta = exp_stats['post_mean'] - exp_stats['pre_mean']
    delta_delta = exp_delta - baseline_delta
    
    return {
        'baseline': baseline_stats,
        'experiment': exp_stats,
        'baseline_delta': baseline_delta,
        'exp_delta': exp_delta,
        'delta_delta': delta_delta
    }

def main():
    # Directory
    directory = r"C:\Users\user\ambs"
    
    # Load data
    print("Loading data...")
    baseline = load_grouped_data(os.path.join(directory, "baseline_grouped.xlsx"))
    experiment = load_grouped_data(os.path.join(directory, "exp_grouped.xlsx"))
    
    print(f"Baseline groups: {list(baseline.keys())}")
    print(f"Experiment groups: {list(experiment.keys())}")
    
    # Get common groups
    groups = sorted(set(baseline.keys()) & set(experiment.keys()))
    print(f"Common groups for comparison: {groups}")
    
    # Key measurement columns to analyze
    key_columns = ['f', 'TVb', 'MVb', 'Penh', 'Ti', 'Te', 'PIFb', 'PEFb']
    
    # Create output directory for plots
    plots_dir = os.path.join(directory, "analysis_plots_v2")
    os.makedirs(plots_dir, exist_ok=True)
    
    # ========================================================================
    # PART 1: PRE vs POST Treatment Analysis for Each Condition
    # ========================================================================
    print("\n" + "="*70)
    print("PART 1: PRE-TREATMENT vs POST-TREATMENT Analysis")
    print("="*70)
    
    pre_post_results = {'baseline': {}, 'experiment': {}}
    
    for condition_name, condition_data in [('baseline', baseline), ('experiment', experiment)]:
        print(f"\n{'='*50}")
        print(f"  {condition_name.upper()}")
        print(f"{'='*50}")
        
        for group in groups:
            print(f"\n--- {group.upper()} ---")
            df = condition_data[group]
            pre_post_results[condition_name][group] = {}
            
            for col in key_columns:
                if col in df.columns:
                    result = calculate_pre_post_stats(df, col)
                    if result:
                        pre_post_results[condition_name][group][col] = result
                        sig = "***" if result['p_value'] and result['p_value'] < 0.001 else \
                              "**" if result['p_value'] and result['p_value'] < 0.01 else \
                              "*" if result['p_value'] and result['p_value'] < 0.05 else ""
                        p_str = f"{result['p_value']:.4f}" if result['p_value'] else 'N/A'
                        print(f"  {col}: Pre={result['pre_mean']:.3f} -> Post={result['post_mean']:.3f}, "
                              f"Change={result['pct_change']:+.1f}%, p={p_str} {sig}")
    
    # ========================================================================
    # PART 2: Treatment Effect Analysis (Delta-Delta)
    # ========================================================================
    print("\n" + "="*70)
    print("PART 2: TREATMENT EFFECT Analysis (Experiment vs Baseline)")
    print("="*70)
    print("(Comparing how pre-to-post changes differ between conditions)")
    
    treatment_effects = {}
    
    for group in groups:
        print(f"\n--- {group.upper()} ---")
        treatment_effects[group] = {}
        
        for col in key_columns:
            result = calculate_treatment_effect(baseline[group], experiment[group], col)
            if result:
                treatment_effects[group][col] = result
                print(f"  {col}:")
                print(f"    Baseline: Pre->Post change = {result['baseline_delta']:+.3f} ({result['baseline']['pct_change']:+.1f}%)")
                print(f"    Experiment: Pre->Post change = {result['exp_delta']:+.3f} ({result['experiment']['pct_change']:+.1f}%)")
                print(f"    Delta-Delta (Exp effect beyond Baseline) = {result['delta_delta']:+.3f}")
    
    # ========================================================================
    # VISUALIZATION 1: Pre vs Post Bar Charts for Each Group
    # ========================================================================
    print("\n\nCreating pre/post comparison visualizations...")
    
    for group in groups:
        fig, axes = plt.subplots(2, 4, figsize=(16, 10))
        axes = axes.flatten()
        
        for idx, col in enumerate(key_columns):
            ax = axes[idx]
            
            x = np.arange(2)  # Baseline, Experiment
            width = 0.35
            
            # Get pre and post values for baseline
            baseline_pre = pre_post_results['baseline'][group].get(col, {}).get('pre_mean', 0)
            baseline_post = pre_post_results['baseline'][group].get(col, {}).get('post_mean', 0)
            baseline_pre_sem = pre_post_results['baseline'][group].get(col, {}).get('pre_std', 0) / np.sqrt(pre_post_results['baseline'][group].get(col, {}).get('pre_n', 1))
            baseline_post_sem = pre_post_results['baseline'][group].get(col, {}).get('post_std', 0) / np.sqrt(pre_post_results['baseline'][group].get(col, {}).get('post_n', 1))
            
            # Get pre and post values for experiment
            exp_pre = pre_post_results['experiment'][group].get(col, {}).get('pre_mean', 0)
            exp_post = pre_post_results['experiment'][group].get(col, {}).get('post_mean', 0)
            exp_pre_sem = pre_post_results['experiment'][group].get(col, {}).get('pre_std', 0) / np.sqrt(pre_post_results['experiment'][group].get(col, {}).get('pre_n', 1))
            exp_post_sem = pre_post_results['experiment'][group].get(col, {}).get('post_std', 0) / np.sqrt(pre_post_results['experiment'][group].get(col, {}).get('post_n', 1))
            
            # Plot bars
            bars1 = ax.bar(x - width/2, [baseline_pre, exp_pre], width, yerr=[baseline_pre_sem, exp_pre_sem],
                          label='Pre-Treatment', color='lightblue', capsize=3, edgecolor='steelblue', linewidth=1.5)
            bars2 = ax.bar(x + width/2, [baseline_post, exp_post], width, yerr=[baseline_post_sem, exp_post_sem],
                          label='Post-Treatment', color='lightcoral', capsize=3, edgecolor='darkred', linewidth=1.5)
            
            # Add significance markers
            for i, (pre_val, post_val, condition) in enumerate([(baseline_pre, baseline_post, 'baseline'), 
                                                                  (exp_pre, exp_post, 'experiment')]):
                if col in pre_post_results[condition][group]:
                    p_val = pre_post_results[condition][group][col].get('p_value')
                    if p_val and p_val < 0.05:
                        max_val = max(pre_val, post_val)
                        stars = "***" if p_val < 0.001 else "**" if p_val < 0.01 else "*"
                        ax.text(i, max_val * 1.1, stars, ha='center', fontsize=10, fontweight='bold')
            
            ax.set_ylabel(col)
            ax.set_title(f'{col}')
            ax.set_xticks(x)
            ax.set_xticklabels(['Baseline', 'Experiment'])
            if idx == 0:
                ax.legend(loc='best', fontsize=8)
        
        plt.suptitle(f'{group.upper()}: Pre-Treatment vs Post-Treatment Comparison', fontsize=14, fontweight='bold')
        plt.tight_layout()
        plt.savefig(os.path.join(plots_dir, f'pre_post_{group.replace(" ", "_")}.png'), dpi=150, bbox_inches='tight')
        plt.close()
        print(f"  Saved: pre_post_{group.replace(' ', '_')}.png")
    
    # ========================================================================
    # VISUALIZATION 2: Time Course with Pre/Post Shading
    # ========================================================================
    print("\nCreating time-course with treatment phases...")
    
    fig, axes = plt.subplots(len(groups), 4, figsize=(18, 4*len(groups)))
    if len(groups) == 1:
        axes = axes.reshape(1, -1)
    
    selected_cols = ['f', 'TVb', 'Penh', 'MVb']
    
    for i, group in enumerate(groups):
        for j, col in enumerate(selected_cols):
            ax = axes[i, j]
            
            for data, name, color, marker in [(baseline[group], 'Baseline', 'steelblue', 'o'), 
                                               (experiment[group], 'Experiment', 'coral', 's')]:
                if col in data.columns and 'Minutes_from_Time0' in data.columns:
                    time_vals = data['Minutes_from_Time0'].values
                    col_vals = data[col].values
                    
                    ax.plot(time_vals, col_vals, f'-{marker}', color=color, 
                           label=name, markersize=4, alpha=0.8, linewidth=1.5)
            
            # Add shading for pre/post treatment
            ax.axvspan(ax.get_xlim()[0], 0, alpha=0.1, color='green', label='Pre-Treatment' if i==0 and j==0 else '')
            ax.axvspan(0, ax.get_xlim()[1], alpha=0.1, color='red', label='Post-Treatment' if i==0 and j==0 else '')
            ax.axvline(x=0, color='black', linestyle='--', linewidth=2, label='Treatment' if i==0 and j==0 else '')
            
            ax.set_xlabel('Minutes from Time 0')
            ax.set_ylabel(col)
            ax.set_title(f'{group.upper()}: {col}')
            
            if i == 0 and j == 0:
                ax.legend(loc='best', fontsize=7)
    
    plt.tight_layout()
    plt.savefig(os.path.join(plots_dir, 'timecourse_with_phases.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print(f"  Saved: timecourse_with_phases.png")
    
    # ========================================================================
    # VISUALIZATION 3: Delta-Delta Heatmap (Treatment Effect)
    # ========================================================================
    print("\nCreating treatment effect heatmap...")
    
    # Create matrices for visualization
    baseline_pct_changes = []
    exp_pct_changes = []
    delta_deltas = []
    
    for group in groups:
        baseline_row = []
        exp_row = []
        dd_row = []
        for col in key_columns:
            if col in treatment_effects[group]:
                baseline_row.append(treatment_effects[group][col]['baseline']['pct_change'])
                exp_row.append(treatment_effects[group][col]['experiment']['pct_change'])
                dd_row.append(treatment_effects[group][col]['delta_delta'])
            else:
                baseline_row.append(0)
                exp_row.append(0)
                dd_row.append(0)
        baseline_pct_changes.append(baseline_row)
        exp_pct_changes.append(exp_row)
        delta_deltas.append(dd_row)
    
    fig, axes = plt.subplots(1, 3, figsize=(18, 5))
    
    # Baseline Pre->Post change
    im1 = axes[0].imshow(baseline_pct_changes, cmap='RdBu_r', aspect='auto', vmin=-100, vmax=100)
    axes[0].set_xticks(np.arange(len(key_columns)))
    axes[0].set_yticks(np.arange(len(groups)))
    axes[0].set_xticklabels(key_columns, rotation=45, ha='right')
    axes[0].set_yticklabels([g.replace('group ', '').upper() for g in groups])
    axes[0].set_title('BASELINE: Pre->Post Change (%)', fontweight='bold')
    for i in range(len(groups)):
        for j in range(len(key_columns)):
            val = baseline_pct_changes[i][j]
            color = 'white' if abs(val) > 50 else 'black'
            axes[0].text(j, i, f'{val:+.0f}%', ha='center', va='center', color=color, fontsize=8)
    plt.colorbar(im1, ax=axes[0], label='% Change')
    
    # Experiment Pre->Post change
    im2 = axes[1].imshow(exp_pct_changes, cmap='RdBu_r', aspect='auto', vmin=-100, vmax=100)
    axes[1].set_xticks(np.arange(len(key_columns)))
    axes[1].set_yticks(np.arange(len(groups)))
    axes[1].set_xticklabels(key_columns, rotation=45, ha='right')
    axes[1].set_yticklabels([g.replace('group ', '').upper() for g in groups])
    axes[1].set_title('EXPERIMENT: Pre->Post Change (%)', fontweight='bold')
    for i in range(len(groups)):
        for j in range(len(key_columns)):
            val = exp_pct_changes[i][j]
            color = 'white' if abs(val) > 50 else 'black'
            axes[1].text(j, i, f'{val:+.0f}%', ha='center', va='center', color=color, fontsize=8)
    plt.colorbar(im2, ax=axes[1], label='% Change')
    
    # Delta-Delta (absolute values)
    delta_deltas_arr = np.array(delta_deltas)
    max_abs = np.nanmax(np.abs(delta_deltas_arr)) if not np.all(np.isnan(delta_deltas_arr)) else 1
    im3 = axes[2].imshow(delta_deltas_arr, cmap='PuOr', aspect='auto', vmin=-max_abs, vmax=max_abs)
    axes[2].set_xticks(np.arange(len(key_columns)))
    axes[2].set_yticks(np.arange(len(groups)))
    axes[2].set_xticklabels(key_columns, rotation=45, ha='right')
    axes[2].set_yticklabels([g.replace('group ', '').upper() for g in groups])
    axes[2].set_title('TREATMENT EFFECT\n(Exp change - Baseline change)', fontweight='bold')
    for i in range(len(groups)):
        for j in range(len(key_columns)):
            val = delta_deltas[i][j]
            color = 'white' if abs(val) > max_abs*0.5 else 'black'
            axes[2].text(j, i, f'{val:+.1f}', ha='center', va='center', color=color, fontsize=8)
    plt.colorbar(im3, ax=axes[2], label='Delta-Delta')
    
    plt.tight_layout()
    plt.savefig(os.path.join(plots_dir, 'treatment_effect_heatmap.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print(f"  Saved: treatment_effect_heatmap.png")
    
    # ========================================================================
    # VISUALIZATION 4: Paired Pre-Post Comparison (Lines)
    # ========================================================================
    print("\nCreating paired pre-post line plots...")
    
    fig, axes = plt.subplots(len(groups), 4, figsize=(16, 4*len(groups)))
    if len(groups) == 1:
        axes = axes.reshape(1, -1)
    
    selected_cols = ['f', 'TVb', 'Penh', 'MVb']
    
    for i, group in enumerate(groups):
        for j, col in enumerate(selected_cols):
            ax = axes[i, j]
            
            # Get data
            baseline_pre = pre_post_results['baseline'][group].get(col, {}).get('pre_mean', np.nan)
            baseline_post = pre_post_results['baseline'][group].get(col, {}).get('post_mean', np.nan)
            exp_pre = pre_post_results['experiment'][group].get(col, {}).get('pre_mean', np.nan)
            exp_post = pre_post_results['experiment'][group].get(col, {}).get('post_mean', np.nan)
            
            # Plot lines
            ax.plot([0, 1], [baseline_pre, baseline_post], 'o-', color='steelblue', 
                   linewidth=2, markersize=10, label='Baseline')
            ax.plot([0, 1], [exp_pre, exp_post], 's-', color='coral', 
                   linewidth=2, markersize=10, label='Experiment')
            
            # Add percentage change annotations
            if not np.isnan(baseline_pre) and not np.isnan(baseline_post) and baseline_pre != 0:
                baseline_pct = ((baseline_post - baseline_pre) / abs(baseline_pre)) * 100
                mid_y = (baseline_pre + baseline_post) / 2
                ax.annotate(f'{baseline_pct:+.1f}%', xy=(0.5, mid_y), 
                           color='steelblue', fontsize=9, fontweight='bold')
            
            if not np.isnan(exp_pre) and not np.isnan(exp_post) and exp_pre != 0:
                exp_pct = ((exp_post - exp_pre) / abs(exp_pre)) * 100
                mid_y = (exp_pre + exp_post) / 2
                ax.annotate(f'{exp_pct:+.1f}%', xy=(0.5, mid_y), 
                           color='coral', fontsize=9, fontweight='bold', 
                           xytext=(0.6, mid_y))
            
            ax.set_xticks([0, 1])
            ax.set_xticklabels(['Pre-Treatment', 'Post-Treatment'])
            ax.set_ylabel(col)
            ax.set_title(f'{group.upper()}: {col}')
            ax.set_xlim(-0.2, 1.4)
            
            if i == 0 and j == 0:
                ax.legend(loc='best', fontsize=8)
    
    plt.tight_layout()
    plt.savefig(os.path.join(plots_dir, 'paired_pre_post_lines.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print(f"  Saved: paired_pre_post_lines.png")
    
    # ========================================================================
    # VISUALIZATION 5: Post-Treatment Time Course (Zoomed)
    # ========================================================================
    print("\nCreating post-treatment time course...")
    
    fig, axes = plt.subplots(len(groups), 4, figsize=(18, 4*len(groups)))
    if len(groups) == 1:
        axes = axes.reshape(1, -1)
    
    for i, group in enumerate(groups):
        for j, col in enumerate(selected_cols):
            ax = axes[i, j]
            
            for data, name, color, marker in [(baseline[group], 'Baseline', 'steelblue', 'o'), 
                                               (experiment[group], 'Experiment', 'coral', 's')]:
                if col in data.columns and 'Minutes_from_Time0' in data.columns:
                    # Only post-treatment data
                    post_data = data[data['Minutes_from_Time0'] > 0]
                    
                    if len(post_data) > 0:
                        time_vals = post_data['Minutes_from_Time0'].values
                        col_vals = post_data[col].values
                        
                        ax.plot(time_vals, col_vals, f'-{marker}', color=color, 
                               label=name, markersize=5, alpha=0.8, linewidth=2)
                        
                        # Add trend line
                        if len(time_vals) > 2:
                            z = np.polyfit(time_vals, col_vals, 1)
                            p = np.poly1d(z)
                            ax.plot(time_vals, p(time_vals), '--', color=color, alpha=0.5, linewidth=1)
            
            ax.set_xlabel('Minutes After Treatment')
            ax.set_ylabel(col)
            ax.set_title(f'{group.upper()}: {col} (Post-Treatment)')
            
            if i == 0 and j == 0:
                ax.legend(loc='best', fontsize=8)
    
    plt.tight_layout()
    plt.savefig(os.path.join(plots_dir, 'post_treatment_timecourse.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print(f"  Saved: post_treatment_timecourse.png")
    
    # ========================================================================
    # VISUALIZATION 6: Recovery Analysis (Normalized to Pre-Treatment)
    # ========================================================================
    print("\nCreating normalized recovery analysis...")
    
    fig, axes = plt.subplots(len(groups), 4, figsize=(18, 4*len(groups)))
    if len(groups) == 1:
        axes = axes.reshape(1, -1)
    
    for i, group in enumerate(groups):
        for j, col in enumerate(selected_cols):
            ax = axes[i, j]
            
            for data, name, color, marker in [(baseline[group], 'Baseline', 'steelblue', 'o'), 
                                               (experiment[group], 'Experiment', 'coral', 's')]:
                if col in data.columns and 'Minutes_from_Time0' in data.columns:
                    # Get pre-treatment mean for normalization
                    pre_data = data[data['Minutes_from_Time0'] < 0]
                    pre_mean = pre_data[col].mean() if len(pre_data) > 0 else 1
                    
                    if pre_mean != 0:
                        # Normalize all data to pre-treatment baseline (100%)
                        time_vals = data['Minutes_from_Time0'].values
                        normalized_vals = (data[col].values / pre_mean) * 100
                        
                        ax.plot(time_vals, normalized_vals, f'-{marker}', color=color, 
                               label=name, markersize=4, alpha=0.8, linewidth=1.5)
            
            # Add reference line at 100%
            ax.axhline(y=100, color='gray', linestyle='--', alpha=0.5, label='Pre-Treatment Level')
            ax.axvline(x=0, color='black', linestyle='-', linewidth=2, alpha=0.5)
            
            ax.set_xlabel('Minutes from Time 0')
            ax.set_ylabel(f'{col} (% of Pre-Treatment)')
            ax.set_title(f'{group.upper()}: {col} Normalized')
            
            if i == 0 and j == 0:
                ax.legend(loc='best', fontsize=7)
    
    plt.tight_layout()
    plt.savefig(os.path.join(plots_dir, 'normalized_recovery.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print(f"  Saved: normalized_recovery.png")
    
    # ========================================================================
    # VISUALIZATION 7: Summary Comparison Bar Chart
    # ========================================================================
    print("\nCreating summary comparison chart...")
    
    fig, axes = plt.subplots(2, 4, figsize=(18, 10))
    axes = axes.flatten()
    
    for idx, col in enumerate(key_columns):
        ax = axes[idx]
        
        x = np.arange(len(groups))
        width = 0.2
        
        baseline_pre_vals = [pre_post_results['baseline'][g].get(col, {}).get('pre_mean', 0) for g in groups]
        baseline_post_vals = [pre_post_results['baseline'][g].get(col, {}).get('post_mean', 0) for g in groups]
        exp_pre_vals = [pre_post_results['experiment'][g].get(col, {}).get('pre_mean', 0) for g in groups]
        exp_post_vals = [pre_post_results['experiment'][g].get(col, {}).get('post_mean', 0) for g in groups]
        
        ax.bar(x - 1.5*width, baseline_pre_vals, width, label='Baseline Pre', color='lightblue', edgecolor='steelblue')
        ax.bar(x - 0.5*width, baseline_post_vals, width, label='Baseline Post', color='steelblue')
        ax.bar(x + 0.5*width, exp_pre_vals, width, label='Exp Pre', color='lightsalmon', edgecolor='coral')
        ax.bar(x + 1.5*width, exp_post_vals, width, label='Exp Post', color='coral')
        
        ax.set_ylabel(col)
        ax.set_title(f'{col} Comparison')
        ax.set_xticks(x)
        ax.set_xticklabels([g.replace('group ', '').upper() for g in groups])
        
        if idx == 0:
            ax.legend(loc='best', fontsize=7)
    
    plt.tight_layout()
    plt.savefig(os.path.join(plots_dir, 'summary_4way_comparison.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print(f"  Saved: summary_4way_comparison.png")
    
    # ========================================================================
    # STATISTICAL ANALYSIS: Two-way comparison
    # ========================================================================
    print("\n" + "="*70)
    print("PART 3: DETAILED STATISTICAL ANALYSIS")
    print("="*70)
    
    print("\n" + "-"*70)
    print("A. PRE-TREATMENT COMPARISON (Baseline vs Experiment)")
    print("-"*70)
    print("Testing if pre-treatment values differ between conditions")
    
    for group in groups:
        print(f"\n{group.upper()}:")
        for col in key_columns:
            baseline_pre, _ = get_pre_post_data(baseline[group])
            exp_pre, _ = get_pre_post_data(experiment[group])
            
            if baseline_pre is not None and exp_pre is not None:
                if col in baseline_pre.columns and col in exp_pre.columns:
                    b_vals = baseline_pre[col].dropna()
                    e_vals = exp_pre[col].dropna()
                    
                    if len(b_vals) > 0 and len(e_vals) > 0:
                        t_stat, p_val = stats.ttest_ind(b_vals, e_vals)
                        sig = "***" if p_val < 0.001 else "**" if p_val < 0.01 else "*" if p_val < 0.05 else ""
                        diff = ((e_vals.mean() - b_vals.mean()) / b_vals.mean() * 100) if b_vals.mean() != 0 else 0
                        print(f"  {col}: Baseline={b_vals.mean():.2f}, Exp={e_vals.mean():.2f}, Diff={diff:+.1f}%, p={p_val:.4f} {sig}")
    
    print("\n" + "-"*70)
    print("B. POST-TREATMENT COMPARISON (Baseline vs Experiment)")
    print("-"*70)
    print("Testing if post-treatment values differ between conditions")
    
    for group in groups:
        print(f"\n{group.upper()}:")
        for col in key_columns:
            _, baseline_post = get_pre_post_data(baseline[group])
            _, exp_post = get_pre_post_data(experiment[group])
            
            if baseline_post is not None and exp_post is not None:
                if col in baseline_post.columns and col in exp_post.columns:
                    b_vals = baseline_post[col].dropna()
                    e_vals = exp_post[col].dropna()
                    
                    if len(b_vals) > 0 and len(e_vals) > 0:
                        t_stat, p_val = stats.ttest_ind(b_vals, e_vals)
                        sig = "***" if p_val < 0.001 else "**" if p_val < 0.01 else "*" if p_val < 0.05 else ""
                        diff = ((e_vals.mean() - b_vals.mean()) / b_vals.mean() * 100) if b_vals.mean() != 0 else 0
                        print(f"  {col}: Baseline={b_vals.mean():.2f}, Exp={e_vals.mean():.2f}, Diff={diff:+.1f}%, p={p_val:.4f} {sig}")
    
    # ========================================================================
    # SUMMARY AND INSIGHTS
    # ========================================================================
    print("\n" + "="*70)
    print("SUMMARY AND KEY INSIGHTS")
    print("="*70)
    
    print("\n[1] PRE-TREATMENT TO POST-TREATMENT CHANGES")
    print("-" * 50)
    
    significant_pre_post = []
    for condition in ['baseline', 'experiment']:
        for group in groups:
            for col in key_columns:
                if col in pre_post_results[condition][group]:
                    result = pre_post_results[condition][group][col]
                    if result['p_value'] and result['p_value'] < 0.05:
                        significant_pre_post.append({
                            'condition': condition,
                            'group': group,
                            'parameter': col,
                            'pct_change': result['pct_change'],
                            'p_value': result['p_value']
                        })
    
    print(f"\nFound {len(significant_pre_post)} significant pre->post changes (p < 0.05):\n")
    
    # Group by condition
    for condition in ['baseline', 'experiment']:
        cond_findings = [f for f in significant_pre_post if f['condition'] == condition]
        if cond_findings:
            print(f"  {condition.upper()}:")
            for finding in sorted(cond_findings, key=lambda x: x['p_value']):
                direction = "INCREASED" if finding['pct_change'] > 0 else "DECREASED"
                print(f"    - {finding['group'].upper()} {finding['parameter']}: "
                      f"{direction} by {abs(finding['pct_change']):.1f}% (p={finding['p_value']:.4f})")
    
    print("\n\n[2] KEY TREATMENT EFFECTS (Delta-Delta Analysis)")
    print("-" * 50)
    print("(Positive = Experiment shows more increase/less decrease than Baseline)")
    print("(Negative = Experiment shows more decrease/less increase than Baseline)\n")
    
    for group in groups:
        print(f"{group.upper()}:")
        for col in key_columns:
            if col in treatment_effects[group]:
                dd = treatment_effects[group][col]['delta_delta']
                bl_change = treatment_effects[group][col]['baseline']['pct_change']
                exp_change = treatment_effects[group][col]['experiment']['pct_change']
                
                if abs(exp_change - bl_change) > 20:  # Only show meaningful differences
                    interpretation = "Larger effect in Experiment" if dd > 0 else "Larger effect in Baseline"
                    print(f"  {col}: Baseline {bl_change:+.1f}% vs Experiment {exp_change:+.1f}% -> {interpretation}")
    
    print("\n\n[3] OVERALL CONCLUSIONS")
    print("-" * 50)
    
    # Analyze Penh (key respiratory indicator)
    print("\nPenh (Airway Resistance Indicator):")
    for group in groups:
        if 'Penh' in treatment_effects[group]:
            bl = treatment_effects[group]['Penh']['baseline']['pct_change']
            exp = treatment_effects[group]['Penh']['experiment']['pct_change']
            print(f"  {group.upper()}: Baseline {bl:+.1f}% vs Experiment {exp:+.1f}%")
            if exp > bl + 50:
                print(f"    -> EXPERIMENT shows SIGNIFICANTLY GREATER Penh increase (airways more constricted)")
    
    # Analyze breathing pattern
    print("\nBreathing Pattern (f = respiratory rate):")
    for group in groups:
        if 'f' in treatment_effects[group]:
            bl = treatment_effects[group]['f']['baseline']['pct_change']
            exp = treatment_effects[group]['f']['experiment']['pct_change']
            print(f"  {group.upper()}: Baseline {bl:+.1f}% vs Experiment {exp:+.1f}%")
    
    print("\n" + "="*70)
    print(f"All visualizations saved to: {plots_dir}")
    print("="*70)
    
    # Save detailed summary to file
    summary_path = os.path.join(directory, "analysis_summary_v2.txt")
    with open(summary_path, 'w') as f:
        f.write("ENHANCED EXPERIMENT ANALYSIS SUMMARY\n")
        f.write("(Pre-Treatment vs Post-Treatment Analysis)\n")
        f.write("="*70 + "\n\n")
        
        f.write("PRE->POST TREATMENT CHANGES\n")
        f.write("-"*50 + "\n\n")
        
        for condition in ['BASELINE', 'EXPERIMENT']:
            f.write(f"{condition}:\n")
            for group in groups:
                f.write(f"\n  {group.upper()}:\n")
                for col in key_columns:
                    if col in pre_post_results[condition.lower()][group]:
                        r = pre_post_results[condition.lower()][group][col]
                        sig = "***" if r['p_value'] and r['p_value'] < 0.001 else \
                              "**" if r['p_value'] and r['p_value'] < 0.01 else \
                              "*" if r['p_value'] and r['p_value'] < 0.05 else ""
                        f.write(f"    {col}: {r['pre_mean']:.2f} -> {r['post_mean']:.2f} ({r['pct_change']:+.1f}%) {sig}\n")
            f.write("\n")
        
        f.write("\nTREATMENT EFFECTS (Delta-Delta)\n")
        f.write("-"*50 + "\n")
        for group in groups:
            f.write(f"\n{group.upper()}:\n")
            for col in key_columns:
                if col in treatment_effects[group]:
                    te = treatment_effects[group][col]
                    f.write(f"  {col}: Baseline change={te['baseline']['pct_change']:+.1f}%, "
                           f"Exp change={te['experiment']['pct_change']:+.1f}%, "
                           f"Delta-Delta={te['delta_delta']:+.2f}\n")
    
    print(f"\nSummary saved to: {summary_path}")

if __name__ == "__main__":
    main()

