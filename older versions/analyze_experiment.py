import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from scipy import stats
import os

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

def get_numeric_columns(df):
    """Get numeric columns excluding Minutes_from_Time0"""
    numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
    if 'Minutes_from_Time0' in numeric_cols:
        numeric_cols.remove('Minutes_from_Time0')
    return numeric_cols

def calculate_stats(baseline_df, exp_df, column):
    """Calculate statistical comparison between baseline and experiment"""
    baseline_vals = baseline_df[column].dropna()
    exp_vals = exp_df[column].dropna()
    
    if len(baseline_vals) == 0 or len(exp_vals) == 0:
        return None
    
    # Calculate basic stats
    baseline_mean = baseline_vals.mean()
    exp_mean = exp_vals.mean()
    baseline_std = baseline_vals.std()
    exp_std = exp_vals.std()
    
    # Percent change
    if baseline_mean != 0:
        pct_change = ((exp_mean - baseline_mean) / abs(baseline_mean)) * 100
    else:
        pct_change = float('inf') if exp_mean != 0 else 0
    
    # T-test (independent samples)
    try:
        t_stat, p_value = stats.ttest_ind(baseline_vals, exp_vals)
    except:
        t_stat, p_value = None, None
    
    return {
        'baseline_mean': baseline_mean,
        'exp_mean': exp_mean,
        'baseline_std': baseline_std,
        'exp_std': exp_std,
        'pct_change': pct_change,
        't_stat': t_stat,
        'p_value': p_value
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
    
    # Key measurement columns to analyze (respiratory parameters)
    key_columns = ['f', 'TVb', 'MVb', 'Penh', 'Ti', 'Te', 'PIFb', 'PEFb']
    
    # Create output directory for plots
    plots_dir = os.path.join(directory, "analysis_plots")
    os.makedirs(plots_dir, exist_ok=True)
    
    # ========================================================================
    # PART 1: Compare Baseline vs Experiment for each group
    # ========================================================================
    print("\n" + "="*70)
    print("PART 1: BASELINE (Control) vs EXPERIMENT Comparison")
    print("="*70)
    
    comparison_results = {}
    
    for group in groups:
        print(f"\n--- {group.upper()} ---")
        baseline_df = baseline[group]
        exp_df = experiment[group]
        
        comparison_results[group] = {}
        
        for col in key_columns:
            if col in baseline_df.columns and col in exp_df.columns:
                result = calculate_stats(baseline_df, exp_df, col)
                if result:
                    comparison_results[group][col] = result
                    significance = "***" if result['p_value'] and result['p_value'] < 0.001 else \
                                   "**" if result['p_value'] and result['p_value'] < 0.01 else \
                                   "*" if result['p_value'] and result['p_value'] < 0.05 else ""
                    p_str = f"{result['p_value']:.4f}" if result['p_value'] else 'N/A'
                    print(f"  {col}: Baseline={result['baseline_mean']:.3f}±{result['baseline_std']:.3f}, "
                          f"Exp={result['exp_mean']:.3f}±{result['exp_std']:.3f}, "
                          f"Change={result['pct_change']:+.1f}%, p={p_str} {significance}")
    
    # ========================================================================
    # VISUALIZATION 1: Time-course comparison for each group
    # ========================================================================
    print("\n\nCreating time-course visualizations...")
    
    fig, axes = plt.subplots(len(groups), len(key_columns), figsize=(20, 4*len(groups)))
    if len(groups) == 1:
        axes = axes.reshape(1, -1)
    
    for i, group in enumerate(groups):
        baseline_df = baseline[group]
        exp_df = experiment[group]
        
        for j, col in enumerate(key_columns):
            ax = axes[i, j]
            
            if col in baseline_df.columns and col in exp_df.columns:
                # Plot baseline
                baseline_time = baseline_df['Minutes_from_Time0'].values
                baseline_vals = baseline_df[col].values
                ax.plot(baseline_time, baseline_vals, 'b-o', label='Baseline', markersize=3, alpha=0.7)
                
                # Plot experiment
                exp_time = exp_df['Minutes_from_Time0'].values
                exp_vals = exp_df[col].values
                ax.plot(exp_time, exp_vals, 'r-s', label='Experiment', markersize=3, alpha=0.7)
                
                # Add vertical line at time 0
                ax.axvline(x=0, color='green', linestyle='--', alpha=0.5, label='Treatment')
                
                ax.set_xlabel('Minutes from Time 0')
                ax.set_ylabel(col)
                ax.set_title(f'{group.upper()}: {col}')
                if i == 0 and j == 0:
                    ax.legend(loc='best', fontsize=8)
    
    plt.tight_layout()
    plt.savefig(os.path.join(plots_dir, 'timecourse_comparison.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print(f"  Saved: timecourse_comparison.png")
    
    # ========================================================================
    # VISUALIZATION 2: Bar chart comparison (mean ± SEM)
    # ========================================================================
    fig, axes = plt.subplots(2, 4, figsize=(16, 10))
    axes = axes.flatten()
    
    for idx, col in enumerate(key_columns):
        ax = axes[idx]
        
        x = np.arange(len(groups))
        width = 0.35
        
        baseline_means = []
        baseline_sems = []
        exp_means = []
        exp_sems = []
        
        for group in groups:
            if col in comparison_results[group]:
                result = comparison_results[group][col]
                baseline_means.append(result['baseline_mean'])
                baseline_sems.append(result['baseline_std'] / np.sqrt(len(baseline[group])))
                exp_means.append(result['exp_mean'])
                exp_sems.append(result['exp_std'] / np.sqrt(len(experiment[group])))
            else:
                baseline_means.append(0)
                baseline_sems.append(0)
                exp_means.append(0)
                exp_sems.append(0)
        
        bars1 = ax.bar(x - width/2, baseline_means, width, yerr=baseline_sems, 
                       label='Baseline', color='steelblue', capsize=3, alpha=0.8)
        bars2 = ax.bar(x + width/2, exp_means, width, yerr=exp_sems,
                       label='Experiment', color='coral', capsize=3, alpha=0.8)
        
        # Add significance markers
        for i, group in enumerate(groups):
            if col in comparison_results[group]:
                p_val = comparison_results[group][col]['p_value']
                if p_val and p_val < 0.05:
                    max_val = max(baseline_means[i], exp_means[i])
                    max_sem = max(baseline_sems[i], exp_sems[i])
                    stars = "***" if p_val < 0.001 else "**" if p_val < 0.01 else "*"
                    ax.text(i, max_val + max_sem * 1.5, stars, ha='center', fontsize=12, fontweight='bold')
        
        ax.set_ylabel(col)
        ax.set_title(f'{col} Comparison')
        ax.set_xticks(x)
        ax.set_xticklabels([g.replace('group ', '').upper() for g in groups])
        ax.legend(loc='best', fontsize=8)
    
    plt.tight_layout()
    plt.savefig(os.path.join(plots_dir, 'bar_comparison.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print(f"  Saved: bar_comparison.png")
    
    # ========================================================================
    # PART 2: Compare between groups within the experiment
    # ========================================================================
    print("\n" + "="*70)
    print("PART 2: BETWEEN-GROUP Comparison (Within Experiment)")
    print("="*70)
    
    # Calculate means for each group in experiment
    exp_group_stats = {}
    for group in groups:
        exp_df = experiment[group]
        exp_group_stats[group] = {}
        for col in key_columns:
            if col in exp_df.columns:
                vals = exp_df[col].dropna()
                exp_group_stats[group][col] = {
                    'mean': vals.mean(),
                    'std': vals.std(),
                    'n': len(vals)
                }
    
    print("\nExperiment Group Statistics:")
    print("-" * 60)
    header = f"{'Parameter':<10}"
    for group in groups:
        header += f" | {group.replace('group ', '').upper():^20}"
    print(header)
    print("-" * 60)
    
    for col in key_columns:
        row = f"{col:<10}"
        for group in groups:
            if col in exp_group_stats[group]:
                mean = exp_group_stats[group][col]['mean']
                std = exp_group_stats[group][col]['std']
                row += f" | {mean:>8.2f} ± {std:<7.2f}"
            else:
                row += f" | {'N/A':^20}"
        print(row)
    
    # ANOVA for each parameter across groups
    print("\n\nOne-way ANOVA (comparing groups within experiment):")
    print("-" * 50)
    anova_results = {}
    for col in key_columns:
        group_values = []
        for group in groups:
            if col in experiment[group].columns:
                group_values.append(experiment[group][col].dropna().values)
        
        if len(group_values) >= 2 and all(len(v) > 0 for v in group_values):
            try:
                f_stat, p_value = stats.f_oneway(*group_values)
                anova_results[col] = {'f_stat': f_stat, 'p_value': p_value}
                significance = "***" if p_value < 0.001 else "**" if p_value < 0.01 else "*" if p_value < 0.05 else ""
                print(f"  {col}: F={f_stat:.3f}, p={p_value:.4f} {significance}")
            except:
                print(f"  {col}: Could not compute ANOVA")
    
    # ========================================================================
    # VISUALIZATION 3: Heatmap of percent changes
    # ========================================================================
    print("\n\nCreating percent change heatmap...")
    
    pct_changes = []
    for group in groups:
        row = []
        for col in key_columns:
            if col in comparison_results[group]:
                row.append(comparison_results[group][col]['pct_change'])
            else:
                row.append(0)
        pct_changes.append(row)
    
    pct_changes = np.array(pct_changes)
    
    fig, ax = plt.subplots(figsize=(12, 6))
    im = ax.imshow(pct_changes, cmap='RdBu_r', aspect='auto', vmin=-100, vmax=100)
    
    ax.set_xticks(np.arange(len(key_columns)))
    ax.set_yticks(np.arange(len(groups)))
    ax.set_xticklabels(key_columns)
    ax.set_yticklabels([g.replace('group ', '').upper() for g in groups])
    
    # Add text annotations
    for i in range(len(groups)):
        for j in range(len(key_columns)):
            val = pct_changes[i, j]
            color = 'white' if abs(val) > 50 else 'black'
            ax.text(j, i, f'{val:+.1f}%', ha='center', va='center', color=color, fontsize=9)
    
    ax.set_title('Percent Change: Experiment vs Baseline (Control)\n(Positive = Increase, Negative = Decrease)', fontsize=12)
    plt.colorbar(im, ax=ax, label='% Change')
    
    plt.tight_layout()
    plt.savefig(os.path.join(plots_dir, 'pct_change_heatmap.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print(f"  Saved: pct_change_heatmap.png")
    
    # ========================================================================
    # VISUALIZATION 4: Box plots for key parameters
    # ========================================================================
    print("\nCreating box plots...")
    
    fig, axes = plt.subplots(2, 4, figsize=(16, 10))
    axes = axes.flatten()
    
    for idx, col in enumerate(key_columns):
        ax = axes[idx]
        
        data_to_plot = []
        labels = []
        colors = []
        
        for group in groups:
            group_letter = group.replace('group ', '').upper()
            
            # Baseline
            if col in baseline[group].columns:
                data_to_plot.append(baseline[group][col].dropna().values)
                labels.append(f'{group_letter}\nBaseline')
                colors.append('steelblue')
            
            # Experiment
            if col in experiment[group].columns:
                data_to_plot.append(experiment[group][col].dropna().values)
                labels.append(f'{group_letter}\nExp')
                colors.append('coral')
        
        bp = ax.boxplot(data_to_plot, labels=labels, patch_artist=True)
        
        for patch, color in zip(bp['boxes'], colors):
            patch.set_facecolor(color)
            patch.set_alpha(0.7)
        
        ax.set_ylabel(col)
        ax.set_title(f'{col} Distribution')
        ax.tick_params(axis='x', rotation=45)
    
    plt.tight_layout()
    plt.savefig(os.path.join(plots_dir, 'boxplots.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print(f"  Saved: boxplots.png")
    
    # ========================================================================
    # VISUALIZATION 5: Pre vs Post treatment comparison
    # ========================================================================
    print("\nCreating pre/post treatment analysis...")
    
    fig, axes = plt.subplots(len(groups), 4, figsize=(16, 4*len(groups)))
    if len(groups) == 1:
        axes = axes.reshape(1, -1)
    
    selected_cols = ['f', 'TVb', 'Penh', 'MVb']
    
    for i, group in enumerate(groups):
        for j, col in enumerate(selected_cols):
            ax = axes[i, j]
            
            for data, name, color in [(baseline[group], 'Baseline', 'steelblue'), 
                                       (experiment[group], 'Experiment', 'coral')]:
                if col in data.columns and 'Minutes_from_Time0' in data.columns:
                    df = data.copy()
                    
                    # Pre-treatment (before time 0)
                    pre = df[df['Minutes_from_Time0'] < 0][col].dropna()
                    # Post-treatment (after time 0, excluding time 0 itself)
                    post = df[df['Minutes_from_Time0'] > 0][col].dropna()
                    
                    if len(pre) > 0 and len(post) > 0:
                        pre_mean = pre.mean()
                        post_mean = post.mean()
                        
                        x_pos = 0 if name == 'Baseline' else 1
                        ax.plot([x_pos-0.1, x_pos+0.1], [pre_mean, post_mean], 
                               'o-', color=color, markersize=10, linewidth=2, 
                               label=f'{name}' if i == 0 and j == 0 else '')
                        
                        # Add arrow to show direction
                        if post_mean > pre_mean:
                            ax.annotate('', xy=(x_pos+0.1, post_mean), xytext=(x_pos-0.1, pre_mean),
                                       arrowprops=dict(arrowstyle='->', color=color, lw=1.5))
            
            ax.set_xticks([0, 1])
            ax.set_xticklabels(['Baseline', 'Experiment'])
            ax.set_ylabel(col)
            ax.set_title(f'{group.replace("group ", "").upper()}: {col} (Pre→Post)')
            
            if i == 0 and j == 0:
                ax.legend(loc='best')
    
    plt.tight_layout()
    plt.savefig(os.path.join(plots_dir, 'pre_post_treatment.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print(f"  Saved: pre_post_treatment.png")
    
    # ========================================================================
    # SUMMARY AND INSIGHTS
    # ========================================================================
    print("\n" + "="*70)
    print("SUMMARY AND KEY INSIGHTS")
    print("="*70)
    
    print("\n[STATISTICAL SUMMARY]:")
    print("-" * 50)
    
    significant_findings = []
    for group in groups:
        for col in key_columns:
            if col in comparison_results[group]:
                result = comparison_results[group][col]
                if result['p_value'] and result['p_value'] < 0.05:
                    significant_findings.append({
                        'group': group,
                        'parameter': col,
                        'baseline_mean': result['baseline_mean'],
                        'exp_mean': result['exp_mean'],
                        'pct_change': result['pct_change'],
                        'p_value': result['p_value']
                    })
    
    if significant_findings:
        print(f"\n* Found {len(significant_findings)} statistically significant differences (p < 0.05):\n")
        for finding in sorted(significant_findings, key=lambda x: x['p_value']):
            direction = "INCREASED" if finding['pct_change'] > 0 else "DECREASED"
            print(f"  - {finding['group'].upper()} - {finding['parameter']}: "
                  f"{direction} by {abs(finding['pct_change']):.1f}% (p={finding['p_value']:.4f})")
    else:
        print("\n  No statistically significant differences found at p < 0.05")
    
    print("\n" + "="*70)
    print("KEY OBSERVATIONS:")
    print("="*70)
    
    # Analyze overall trends
    print("\n1. RESPIRATORY FREQUENCY (f):")
    for group in groups:
        if 'f' in comparison_results[group]:
            result = comparison_results[group]['f']
            change = "increased" if result['pct_change'] > 0 else "decreased"
            print(f"   {group.upper()}: {change} by {abs(result['pct_change']):.1f}% "
                  f"(Baseline: {result['baseline_mean']:.2f}, Exp: {result['exp_mean']:.2f})")
    
    print("\n2. TIDAL VOLUME (TVb):")
    for group in groups:
        if 'TVb' in comparison_results[group]:
            result = comparison_results[group]['TVb']
            change = "increased" if result['pct_change'] > 0 else "decreased"
            print(f"   {group.upper()}: {change} by {abs(result['pct_change']):.1f}% "
                  f"(Baseline: {result['baseline_mean']:.3f}, Exp: {result['exp_mean']:.3f})")
    
    print("\n3. ENHANCED PAUSE (Penh) - Airway resistance indicator:")
    for group in groups:
        if 'Penh' in comparison_results[group]:
            result = comparison_results[group]['Penh']
            change = "increased" if result['pct_change'] > 0 else "decreased"
            sig = " *SIGNIFICANT*" if result['p_value'] and result['p_value'] < 0.05 else ""
            print(f"   {group.upper()}: {change} by {abs(result['pct_change']):.1f}%{sig}")
    
    print("\n4. MINUTE VENTILATION (MVb):")
    for group in groups:
        if 'MVb' in comparison_results[group]:
            result = comparison_results[group]['MVb']
            change = "increased" if result['pct_change'] > 0 else "decreased"
            print(f"   {group.upper()}: {change} by {abs(result['pct_change']):.1f}%")
    
    print("\n" + "="*70)
    print(f"All visualizations saved to: {plots_dir}")
    print("="*70)
    
    # Save summary to file
    summary_path = os.path.join(directory, "analysis_summary.txt")
    with open(summary_path, 'w') as f:
        f.write("EXPERIMENT ANALYSIS SUMMARY\n")
        f.write("="*70 + "\n\n")
        f.write("Comparison: Baseline (Control) vs Experiment\n")
        f.write(f"Groups analyzed: {', '.join(groups)}\n\n")
        
        f.write("SIGNIFICANT FINDINGS (p < 0.05):\n")
        f.write("-"*50 + "\n")
        for finding in sorted(significant_findings, key=lambda x: x['p_value']):
            direction = "INCREASED" if finding['pct_change'] > 0 else "DECREASED"
            f.write(f"{finding['group'].upper()} - {finding['parameter']}: "
                   f"{direction} by {abs(finding['pct_change']):.1f}% (p={finding['p_value']:.4f})\n")
        
        f.write("\n\nDETAILED STATISTICS:\n")
        f.write("-"*50 + "\n")
        for group in groups:
            f.write(f"\n{group.upper()}:\n")
            for col in key_columns:
                if col in comparison_results[group]:
                    result = comparison_results[group][col]
                    f.write(f"  {col}: Baseline={result['baseline_mean']:.3f}±{result['baseline_std']:.3f}, "
                           f"Exp={result['exp_mean']:.3f}±{result['exp_std']:.3f}, "
                           f"Change={result['pct_change']:+.1f}%, p={result['p_value']:.4f}\n")
    
    print(f"\nSummary saved to: {summary_path}")

if __name__ == "__main__":
    main()


