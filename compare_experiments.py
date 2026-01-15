import pandas as pd
import matplotlib
matplotlib.use('Agg')  # Use non-interactive backend
import matplotlib.pyplot as plt
import numpy as np
from scipy import stats
from scipy.stats import f_oneway
from itertools import combinations
import os
import gc
import glob
import warnings
from datetime import datetime
warnings.filterwarnings('ignore')

# Set style for better looking plots
plt.style.use('seaborn-v0_8-whitegrid')
plt.rcParams['figure.figsize'] = (16, 10)
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
    """Get baseline data (post-treatment only, as reference)
    Ignores first 2 minutes after treatment - starts from minute 3"""
    if 'Minutes_from_Time0' not in df.columns:
        return df
    return df[df['Minutes_from_Time0'] >= 3].copy()

def main():
    # Directory
    directory = os.getcwd()
    
    # Generate timestamp prefix for file names
    timestamp = datetime.now().strftime("%Y-%m-%d %H%M%S")
    
    # ========================================================================
    # Define experiments to compare
    # Each experiment is a tuple: (name, baseline_file_pattern, experiment_file_pattern)
    # ========================================================================
    experiments = [
        {
            'name': 'Antidote 301225',
            'baseline_pattern': '*antidote baseline 301225_grouped.xlsx',
            'experiment_pattern': '*antidote 301225_grouped.xlsx',
        },
        {
            'name': 'Original Experiment',
            'baseline_pattern': 'baseline_grouped.xlsx',
            'experiment_pattern': 'exp_grouped.xlsx',
        },
    ]
    
    # Colors and markers for experiments and groups
    # Using distinct color palettes for each experiment
    experiment_colors = {
        'Antidote 301225': {
            'group a': '#E74C3C',  # Red
            'group b': '#3498DB',  # Blue
            'group c': '#27AE60',  # Green
            'group d': '#9B59B6',  # Purple
            'group e': '#F39C12',  # Orange
            'group f': '#1ABC9C',  # Teal
        },
        'Original Experiment': {
            'group a': '#C0392B',  # Dark Red
            'group b': '#2980B9',  # Dark Blue
            'group c': '#1E8449',  # Dark Green
            'group d': '#7D3C98',  # Dark Purple
            'group e': '#D68910',  # Dark Orange
            'group f': '#148F77',  # Dark Teal
        },
    }
    
    experiment_markers = {
        'Antidote 301225': 'o',      # Circle
        'Original Experiment': 's',   # Square
    }
    
    experiment_linestyles = {
        'Antidote 301225': '-',       # Solid
        'Original Experiment': '--',   # Dashed
    }
    
    # ========================================================================
    # Load all experiment data
    # ========================================================================
    all_experiment_data = {}
    
    for exp in experiments:
        exp_name = exp['name']
        print(f"\nLoading {exp_name}...")
        
        # Find files
        baseline_files = sorted(glob.glob(os.path.join(directory, exp['baseline_pattern'])), reverse=True)
        experiment_files = sorted(glob.glob(os.path.join(directory, exp['experiment_pattern'])), reverse=True)
        
        if not baseline_files:
            print(f"  WARNING: No baseline file found for pattern: {exp['baseline_pattern']}")
            continue
        if not experiment_files:
            print(f"  WARNING: No experiment file found for pattern: {exp['experiment_pattern']}")
            continue
        
        baseline_file = baseline_files[0]
        experiment_file = experiment_files[0]
        
        print(f"  Baseline: {os.path.basename(baseline_file)}")
        print(f"  Experiment: {os.path.basename(experiment_file)}")
        
        baseline_data = load_grouped_data(baseline_file)
        experiment_data = load_grouped_data(experiment_file)
        
        # Get common groups
        groups = sorted(set(baseline_data.keys()) & set(experiment_data.keys()))
        print(f"  Groups: {groups}")
        
        all_experiment_data[exp_name] = {
            'baseline': baseline_data,
            'experiment': experiment_data,
            'groups': groups,
        }
    
    # ========================================================================
    # Create combined time-course plots (2 figures, 4 plots each)
    # ========================================================================
    print("\n" + "="*70)
    print("Creating combined time-course plots...")
    print("="*70)
    
    selected_cols = ['f', 'TVb', 'MVb', 'Penh', 'Ti', 'Te', 'PIFb', 'PEFb']
    
    # Create output directory
    plots_dir = os.path.join(directory, "comparison_plots")
    os.makedirs(plots_dir, exist_ok=True)
    
    # Split into two figures
    param_sets = [
        (selected_cols[:4], 'combined_timecourse_1.png', 'Time Course Comparison (Part 1): f, TVb, MVb, Penh'),
        (selected_cols[4:], 'combined_timecourse_2.png', 'Time Course Comparison (Part 2): Ti, Te, PIFb, PEFb'),
    ]
    
    for param_set, filename_suffix, figure_title in param_sets:
        fig, axes = plt.subplots(2, 2, figsize=(16, 12))
        axes = axes.flatten()
        
        # Track all legend entries for this figure
        legend_handles = []
        legend_labels = []
        
        for j, col in enumerate(param_set):
            ax = axes[j]
            
            for exp_name, exp_data in all_experiment_data.items():
                baseline_data = exp_data['baseline']
                experiment_data = exp_data['experiment']
                groups = exp_data['groups']
                
                marker = experiment_markers.get(exp_name, 'o')
                linestyle = experiment_linestyles.get(exp_name, '-')
                
                for group in groups:
                    # Get color for this experiment/group combination
                    if exp_name in experiment_colors and group in experiment_colors[exp_name]:
                        color = experiment_colors[exp_name][group]
                    else:
                        # Generate a color if not predefined
                        color = plt.cm.tab20(hash(f"{exp_name}_{group}") % 20)
                    
                    baseline_full = baseline_data.get(group)
                    experiment = experiment_data.get(group)
                    
                    if baseline_full is None or experiment is None:
                        continue
                    
                    all_times = []
                    all_means = []
                    all_sems = []
                    
                    # Get the last 10 minutes of baseline data and map to -20 to -10
                    if 'Minutes_from_Time0' in baseline_full.columns and col in baseline_full.columns:
                        baseline_post = baseline_full[baseline_full['Minutes_from_Time0'] >= 3].copy()
                        if len(baseline_post) > 0:
                            max_time = baseline_post['Minutes_from_Time0'].max()
                            last_10_min = baseline_post[baseline_post['Minutes_from_Time0'] >= max_time - 10].copy()
                            if len(last_10_min) > 0:
                                orig_min = last_10_min['Minutes_from_Time0'].min()
                                orig_max = last_10_min['Minutes_from_Time0'].max()
                                if orig_max > orig_min:
                                    last_10_min['mapped_time'] = -20 + (last_10_min['Minutes_from_Time0'] - orig_min) / (orig_max - orig_min) * 10
                                else:
                                    last_10_min['mapped_time'] = -15
                                
                                last_10_min['mapped_minute'] = last_10_min['mapped_time'].round().astype(int)
                                for minute, grp in last_10_min.groupby('mapped_minute'):
                                    vals = grp[col].dropna()
                                    if len(vals) > 0:
                                        all_times.append(minute)
                                        all_means.append(vals.mean())
                                        all_sems.append(vals.std() / np.sqrt(len(vals)) if len(vals) > 1 else 0)
                    
                    # Get experiment data (limit to 30 min post-treatment)
                    if col in experiment.columns and 'Minutes_from_Time0' in experiment.columns:
                        exp_data_df = experiment[(experiment['Minutes_from_Time0'] <= 30)].copy()
                        exp_data_df['minute'] = exp_data_df['Minutes_from_Time0'].round().astype(int)
                        
                        for minute, grp in exp_data_df.groupby('minute'):
                            vals = grp[col].dropna()
                            if len(vals) > 0:
                                all_times.append(minute)
                                all_means.append(vals.mean())
                                all_sems.append(vals.std() / np.sqrt(len(vals)) if len(vals) > 1 else 0)
                    
                    # Sort by time
                    if all_times:
                        sorted_idx = np.argsort(all_times)
                        all_times = np.array(all_times)[sorted_idx]
                        all_means = np.array(all_means)[sorted_idx]
                        all_sems = np.array(all_sems)[sorted_idx]
                        
                        # Create label
                        label = f"{exp_name} - {group.replace('group ', 'Grp ').upper()}"
                        
                        # Plot with error bars
                        line = ax.errorbar(all_times, all_means, yerr=all_sems, 
                                   fmt=f'{linestyle}{marker}', color=color, 
                                   label=label,
                                   markersize=5, alpha=0.8, linewidth=1.5, capsize=2, capthick=1)
                        
                        # Add to legend only once (for first subplot)
                        if j == 0:
                            legend_handles.append(line)
                            legend_labels.append(label)
            
            # Add vertical lines
            ax.axvline(x=0, color='black', linestyle='--', linewidth=2, label='Treatment' if j == 0 else '')
            ax.axvline(x=-10, color='navy', linestyle='--', linewidth=1.5, alpha=0.7)
            
            # Set x-axis limits
            ax.set_xlim(-25, 32)
            
            # Add shading: blue (-20 to -10), green (-10 to 0), red (0 to 30)
            ax.axvspan(-20, -10, alpha=0.2, color='blue')
            ax.axvspan(-10, 0, alpha=0.2, color='green')
            ax.axvspan(0, 30, alpha=0.2, color='red')
            
            ax.set_xlabel('Minutes from Treatment', fontsize=10)
            ax.set_ylabel(col, fontsize=10)
            ax.set_title(f'{col} - All Experiments Time Course', fontsize=12, fontweight='bold')
        
        # Add legend outside the plots
        fig.legend(legend_handles, legend_labels, 
                   loc='center right', 
                   bbox_to_anchor=(1.18, 0.5),
                   fontsize=9,
                   title='Experiment - Group')
        
        plt.suptitle(f'{figure_title}\n(Solid line = Antidote 301225, Dashed line = Original Experiment)', 
                     fontsize=14, fontweight='bold')
        plt.tight_layout()
        plt.subplots_adjust(right=0.82)  # Make room for legend
        
        filename = f'{timestamp} {filename_suffix}'
        plt.savefig(os.path.join(plots_dir, filename), dpi=150, bbox_inches='tight')
        plt.close()
        gc.collect()
        print(f"  Saved: {filename}")
    
    print(f"Location: {plots_dir}")
    
    # ========================================================================
    # NEW PLOT 1: Combined time-course plot for minutes 2-20 with SEM
    # ========================================================================
    print("\n" + "="*70)
    print("Creating combined time-course plots (minutes 2-20)...")
    print("="*70)
    
    # Collect all groups across all experiments for unified color scheme
    all_groups_unified = []
    for exp_name, exp_data in all_experiment_data.items():
        for group in exp_data['groups']:
            group_label = f"{exp_name} - {group}"
            if group_label not in all_groups_unified:
                all_groups_unified.append(group_label)
    
    # Unified color palette for all groups
    unified_colors = {
        'Antidote 301225 - group d': '#9B59B6',  # Purple
        'Antidote 301225 - group e': '#F39C12',  # Orange
        'Original Experiment - group a': '#E74C3C',  # Red
        'Original Experiment - group b': '#3498DB',  # Blue
        'Original Experiment - group c': '#27AE60',  # Green
    }
    
    unified_markers = {
        'Antidote 301225 - group d': 'o',
        'Antidote 301225 - group e': 's',
        'Original Experiment - group a': '^',
        'Original Experiment - group b': 'D',
        'Original Experiment - group c': 'v',
    }
    
    # Function to remove outliers (points that are 4x higher/lower than SEM from mean)
    def remove_outliers(vals, threshold=4):
        """Remove outliers that are more than threshold * SEM from the mean"""
        if len(vals) < 3:
            return vals
        mean_val = np.mean(vals)
        sem_val = np.std(vals) / np.sqrt(len(vals))
        if sem_val == 0:
            return vals
        # Keep values within threshold * SEM of the mean
        # Using SEM-based threshold: if |value - mean| > threshold * SEM, it's an outlier
        filtered = vals[np.abs(vals - mean_val) <= threshold * sem_val]
        if len(filtered) == 0:
            return vals  # Return original if all would be removed
        return filtered
    
    # Parameters where the first datapoint should be removed for group D
    group_d_skip_first_params = ['TVb', 'MVb', 'PIFb', 'PEFb']
    
    # Split into two figures for minutes 2-20
    param_sets_2_20 = [
        (selected_cols[:4], 'combined_timecourse_2to20_part1.png', 'Time Course (Minutes 2-20, Part 1): f, TVb, MVb, Penh'),
        (selected_cols[4:], 'combined_timecourse_2to20_part2.png', 'Time Course (Minutes 2-20, Part 2): Ti, Te, PIFb, PEFb'),
    ]
    
    outliers_removed_count = 0
    first_points_removed_count = 0
    
    for param_set, filename_suffix, figure_title in param_sets_2_20:
        fig, axes = plt.subplots(2, 2, figsize=(16, 12))
        axes = axes.flatten()
        
        legend_handles = []
        legend_labels = []
        
        for j, col in enumerate(param_set):
            ax = axes[j]
            
            for exp_name, exp_data in all_experiment_data.items():
                experiment_data = exp_data['experiment']
                groups = exp_data['groups']
                
                for group in groups:
                    group_label = f"{exp_name} - {group}"
                    color = unified_colors.get(group_label, 'gray')
                    marker = unified_markers.get(group_label, 'o')
                    
                    experiment = experiment_data.get(group)
                    if experiment is None:
                        continue
                    
                    all_times = []
                    all_means = []
                    all_sems = []
                    
                    # Get experiment data for minutes 2-20
                    if col in experiment.columns and 'Minutes_from_Time0' in experiment.columns:
                        exp_data_df = experiment[
                            (experiment['Minutes_from_Time0'] >= 2) & 
                            (experiment['Minutes_from_Time0'] <= 20)
                        ].copy()
                        exp_data_df['minute'] = exp_data_df['Minutes_from_Time0'].round().astype(int)
                        
                        for minute, grp in exp_data_df.groupby('minute'):
                            vals = grp[col].dropna().values
                            if len(vals) > 0:
                                # Remove outliers (points 4x higher/lower than SEM from mean)
                                original_count = len(vals)
                                vals_filtered = remove_outliers(vals, threshold=4)
                                outliers_removed_count += (original_count - len(vals_filtered))
                                
                                if len(vals_filtered) > 0:
                                    all_times.append(minute)
                                    all_means.append(np.mean(vals_filtered))
                                    all_sems.append(np.std(vals_filtered) / np.sqrt(len(vals_filtered)) if len(vals_filtered) > 1 else 0)
                    
                    # Sort by time and plot
                    if all_times:
                        sorted_idx = np.argsort(all_times)
                        all_times = np.array(all_times)[sorted_idx]
                        all_means = np.array(all_means)[sorted_idx]
                        all_sems = np.array(all_sems)[sorted_idx]
                        
                        # Remove first datapoint for group D in specific parameters
                        if group == 'group d' and col in group_d_skip_first_params and len(all_times) > 1:
                            all_times = all_times[1:]
                            all_means = all_means[1:]
                            all_sems = all_sems[1:]
                            first_points_removed_count += 1
                        
                        label = f"{exp_name.replace('Antidote 301225', 'Antidote').replace('Original Experiment', 'Original')} - {group.replace('group ', 'Grp ').upper()}"
                        
                        line = ax.errorbar(all_times, all_means, yerr=all_sems, 
                                   fmt=f'-{marker}', color=color, 
                                   label=label,
                                   markersize=6, alpha=0.9, linewidth=2, capsize=3, capthick=1)
                        
                        if j == 0:
                            legend_handles.append(line)
                            legend_labels.append(label)
            
            ax.set_xlim(1.5, 20.5)
            ax.set_xlabel('Minutes from Treatment', fontsize=11)
            ax.set_ylabel(col, fontsize=11)
            ax.set_title(f'{col} - All Groups (Minutes 2-20)', fontsize=12, fontweight='bold')
            ax.grid(True, alpha=0.3)
        
        fig.legend(legend_handles, legend_labels, 
                   loc='center right', 
                   bbox_to_anchor=(1.15, 0.5),
                   fontsize=10,
                   title='Experiment - Group')
        
        plt.suptitle(f'{figure_title}\n(Outliers removed: points > 4x SEM from mean)', fontsize=14, fontweight='bold')
        plt.tight_layout()
        plt.subplots_adjust(right=0.85)
        
        filename = f'{timestamp} {filename_suffix}'
        plt.savefig(os.path.join(plots_dir, filename), dpi=150, bbox_inches='tight')
        plt.close()
        gc.collect()
        print(f"  Saved: {filename}")
    
    print(f"  Total outliers removed: {outliers_removed_count}")
    print(f"  First datapoints removed for Group D (TVb, MVb, PIFb, PEFb): {first_points_removed_count}")
    
    # ========================================================================
    # NEW PLOT 2: Specific time points (2.5, 4.8, 5.5, 7.5 min) with ANOVA
    # ========================================================================
    print("\n" + "="*70)
    print("Creating specific time point plots with ANOVA analysis...")
    print("="*70)
    
    target_timepoints = [2.5, 4.8, 5.5, 7.5]
    tolerance = 0.5  # Allow +/- 0.5 minutes tolerance for matching
    
    # Collect data for all groups at specific timepoints
    def get_value_at_timepoint(df, col, target_time, tol=0.5):
        """Get values at a specific timepoint with tolerance"""
        if 'Minutes_from_Time0' not in df.columns or col not in df.columns:
            return np.array([])
        
        mask = (df['Minutes_from_Time0'] >= target_time - tol) & (df['Minutes_from_Time0'] <= target_time + tol)
        vals = df.loc[mask, col].dropna().values
        return vals
    
    # Prepare data structure for all groups at all timepoints
    timepoint_data = {tp: {col: {} for col in selected_cols} for tp in target_timepoints}
    
    # Parameters where the first timepoint (2.5 min) should be skipped for group D
    group_d_skip_first_timepoint_params = ['TVb', 'MVb', 'PIFb', 'PEFb']
    first_timepoint = target_timepoints[0]  # 2.5 minutes
    
    for exp_name, exp_data in all_experiment_data.items():
        experiment_data = exp_data['experiment']
        groups = exp_data['groups']
        
        for group in groups:
            group_label = f"{exp_name} - {group}"
            experiment = experiment_data.get(group)
            
            if experiment is None:
                continue
            
            for tp in target_timepoints:
                for col in selected_cols:
                    # Skip first timepoint (2.5 min) for group D in specific parameters
                    if group == 'group d' and tp == first_timepoint and col in group_d_skip_first_timepoint_params:
                        continue
                    
                    vals = get_value_at_timepoint(experiment, col, tp, tolerance)
                    if len(vals) > 0:
                        timepoint_data[tp][col][group_label] = vals
    
    # Perform ANOVA and post-hoc analysis
    anova_results = []
    posthoc_results = []
    
    for col in selected_cols:
        for tp in target_timepoints:
            groups_at_tp = timepoint_data[tp][col]
            group_names = list(groups_at_tp.keys())
            
            if len(group_names) < 2:
                continue
            
            # Prepare data for ANOVA
            group_data = [groups_at_tp[g] for g in group_names if len(groups_at_tp[g]) > 0]
            group_names_valid = [g for g in group_names if len(groups_at_tp[g]) > 0]
            
            if len(group_data) < 2:
                continue
            
            # Perform one-way ANOVA
            try:
                f_stat, p_value = f_oneway(*group_data)
                anova_results.append({
                    'Parameter': col,
                    'Timepoint': tp,
                    'F_statistic': f_stat,
                    'p_value': p_value,
                    'n_groups': len(group_data),
                    'significant': p_value < 0.05
                })
                
                # Post-hoc pairwise comparisons (t-tests with no correction for simplicity)
                for (g1, g2) in combinations(range(len(group_names_valid)), 2):
                    name1 = group_names_valid[g1]
                    name2 = group_names_valid[g2]
                    data1 = groups_at_tp[name1]
                    data2 = groups_at_tp[name2]
                    
                    if len(data1) > 0 and len(data2) > 0:
                        t_stat, p_val = stats.ttest_ind(data1, data2)
                        posthoc_results.append({
                            'Parameter': col,
                            'Timepoint': tp,
                            'Group1': name1,
                            'Group2': name2,
                            'Mean1': np.mean(data1),
                            'Mean2': np.mean(data2),
                            'SEM1': np.std(data1) / np.sqrt(len(data1)) if len(data1) > 1 else 0,
                            'SEM2': np.std(data2) / np.sqrt(len(data2)) if len(data2) > 1 else 0,
                            't_statistic': t_stat,
                            'p_value': p_val,
                            'significant': p_val < 0.05
                        })
            except Exception as e:
                print(f"  Warning: ANOVA failed for {col} at {tp} min: {e}")
    
    # Save ANOVA results to file
    anova_filename = f'{timestamp} anova_analysis_results.txt'
    anova_filepath = os.path.join(plots_dir, anova_filename)
    
    with open(anova_filepath, 'w', encoding='utf-8') as f:
        f.write("ANOVA AND POST-HOC ANALYSIS RESULTS\n")
        f.write("="*80 + "\n")
        f.write(f"Analysis Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"Time Points Analyzed: {target_timepoints}\n")
        f.write(f"Tolerance: +/- {tolerance} minutes\n")
        f.write("\n" + "="*80 + "\n\n")
        
        f.write("ONE-WAY ANOVA RESULTS\n")
        f.write("-"*80 + "\n")
        f.write(f"{'Parameter':<12} {'Timepoint':<12} {'F-statistic':<15} {'p-value':<15} {'Significant':<12}\n")
        f.write("-"*80 + "\n")
        
        for result in anova_results:
            sig = "***" if result['p_value'] < 0.001 else "**" if result['p_value'] < 0.01 else "*" if result['p_value'] < 0.05 else ""
            f.write(f"{result['Parameter']:<12} {result['Timepoint']:<12.1f} {result['F_statistic']:<15.4f} {result['p_value']:<15.6f} {sig:<12}\n")
        
        f.write("\n\n" + "="*80 + "\n")
        f.write("POST-HOC PAIRWISE COMPARISONS (t-tests)\n")
        f.write("="*80 + "\n\n")
        
        for col in selected_cols:
            f.write(f"\n{'='*80}\n")
            f.write(f"PARAMETER: {col}\n")
            f.write(f"{'='*80}\n\n")
            
            col_results = [r for r in posthoc_results if r['Parameter'] == col]
            
            for tp in target_timepoints:
                tp_results = [r for r in col_results if r['Timepoint'] == tp]
                if not tp_results:
                    continue
                
                f.write(f"\nTimepoint: {tp} minutes\n")
                f.write("-"*70 + "\n")
                
                for r in tp_results:
                    sig = "***" if r['p_value'] < 0.001 else "**" if r['p_value'] < 0.01 else "*" if r['p_value'] < 0.05 else "ns"
                    f.write(f"  {r['Group1'][:30]:<32} vs {r['Group2'][:30]:<32}\n")
                    f.write(f"    Mean1={r['Mean1']:.3f} (SEM={r['SEM1']:.3f}), Mean2={r['Mean2']:.3f} (SEM={r['SEM2']:.3f})\n")
                    f.write(f"    t={r['t_statistic']:.3f}, p={r['p_value']:.6f} {sig}\n\n")
        
        f.write("\n\nSIGNIFICANCE LEVELS:\n")
        f.write("  * p < 0.05\n")
        f.write("  ** p < 0.01\n")
        f.write("  *** p < 0.001\n")
        f.write("  ns = not significant\n")
    
    print(f"  ANOVA results saved to: {anova_filename}")
    
    # ========================================================================
    # Export timepoint comparison data to Excel
    # ========================================================================
    print("\n  Exporting timepoint comparison data to Excel...")
    
    # Create a list to store all data for export
    export_data = []
    
    for col in selected_cols:
        for tp in target_timepoints:
            for group_label, vals in timepoint_data[tp][col].items():
                if len(vals) > 0:
                    export_data.append({
                        'Parameter': col,
                        'Timepoint_min': tp,
                        'Group': group_label,
                        'Mean': np.mean(vals),
                        'SEM': np.std(vals) / np.sqrt(len(vals)) if len(vals) > 1 else 0,
                        'STD': np.std(vals),
                        'N': len(vals),
                        'Min': np.min(vals),
                        'Max': np.max(vals),
                        'Raw_Values': ';'.join([f'{v:.4f}' for v in vals])
                    })
    
    # Create DataFrame and save to Excel
    export_df = pd.DataFrame(export_data)
    excel_filename = f'{timestamp} timepoint_comparison_data.xlsx'
    excel_filepath = os.path.join(plots_dir, excel_filename)
    
    # Create Excel with multiple sheets for easier reading
    with pd.ExcelWriter(excel_filepath, engine='openpyxl') as writer:
        # Full data sheet
        export_df.to_excel(writer, sheet_name='All_Data', index=False)
        
        # Create pivot tables for each parameter
        for col in selected_cols:
            col_data = export_df[export_df['Parameter'] == col].copy()
            if len(col_data) > 0:
                # Create pivot table with means
                pivot_mean = col_data.pivot_table(
                    values='Mean', 
                    index='Group', 
                    columns='Timepoint_min',
                    aggfunc='first'
                )
                # Create pivot table with SEMs
                pivot_sem = col_data.pivot_table(
                    values='SEM', 
                    index='Group', 
                    columns='Timepoint_min',
                    aggfunc='first'
                )
                
                # Combine mean and SEM
                combined = pd.DataFrame()
                for tp in target_timepoints:
                    if tp in pivot_mean.columns:
                        combined[f'{tp}_Mean'] = pivot_mean[tp]
                        combined[f'{tp}_SEM'] = pivot_sem[tp]
                
                combined.to_excel(writer, sheet_name=f'{col}_Summary')
    
    print(f"  Data exported to: {excel_filename}")
    
    # Also save as CSV for easy access
    csv_filename = f'{timestamp} timepoint_comparison_data.csv'
    csv_filepath = os.path.join(plots_dir, csv_filename)
    export_df.to_csv(csv_filepath, index=False)
    print(f"  Data exported to: {csv_filename}")
    
    # Create the specific timepoint plots with horizontal significance lines
    param_sets_timepoints = [
        (selected_cols[:4], 'timepoint_comparison_part1.png', 'Parameter Comparison at Specific Timepoints (Part 1)'),
        (selected_cols[4:], 'timepoint_comparison_part2.png', 'Parameter Comparison at Specific Timepoints (Part 2)'),
    ]
    
    # Helper function to draw significance bracket
    def draw_significance_bracket(ax, x1, x2, y, h, stars):
        """Draw a horizontal bracket between two x positions with stars above"""
        ax.plot([x1, x1, x2, x2], [y, y+h, y+h, y], color='black', linewidth=1.2)
        ax.text((x1+x2)/2, y+h, stars, ha='center', va='bottom', fontsize=10, fontweight='bold', color='black')
    
    for param_set, filename_suffix, figure_title in param_sets_timepoints:
        fig, axes = plt.subplots(2, 2, figsize=(18, 14))
        axes = axes.flatten()
        
        for j, col in enumerate(param_set):
            ax = axes[j]
            
            # Collect all groups that have data for this parameter
            groups_with_data = set()
            for tp in target_timepoints:
                groups_with_data.update(timepoint_data[tp][col].keys())
            groups_with_data = sorted(list(groups_with_data))
            
            if not groups_with_data:
                ax.text(0.5, 0.5, 'No data available', ha='center', va='center', transform=ax.transAxes)
                ax.set_title(f'{col}', fontsize=12, fontweight='bold')
                continue
            
            # X positions for timepoints
            x_positions = np.arange(len(target_timepoints))
            width = 0.12  # Width of each group's offset
            n_groups = len(groups_with_data)
            
            # Store group positions for significance brackets
            group_x_positions = {tp: {} for tp in target_timepoints}
            
            # Plot each group
            for i, group_label in enumerate(groups_with_data):
                color = unified_colors.get(group_label, 'gray')
                marker = unified_markers.get(group_label, 'o')
                
                # Offset for this group
                offset = (i - (n_groups - 1) / 2) * width
                
                means = []
                sems = []
                x_vals = []
                
                for k, tp in enumerate(target_timepoints):
                    if group_label in timepoint_data[tp][col]:
                        vals = timepoint_data[tp][col][group_label]
                        if len(vals) > 0:
                            x_pos = x_positions[k] + offset
                            means.append(np.mean(vals))
                            sems.append(np.std(vals) / np.sqrt(len(vals)) if len(vals) > 1 else 0)
                            x_vals.append(x_pos)
                            group_x_positions[tp][group_label] = x_pos
                
                if means:
                    short_label = group_label.replace('Antidote 301225', 'Ant').replace('Original Experiment', 'Orig').replace('group ', 'Grp ')
                    ax.errorbar(x_vals, means, yerr=sems, 
                               fmt=marker, color=color, 
                               label=short_label,
                               markersize=10, capsize=5, capthick=2, elinewidth=2, 
                               linestyle='none')
            
            # Add significance brackets with horizontal lines
            for k, tp in enumerate(target_timepoints):
                tp_posthoc = [r for r in posthoc_results if r['Parameter'] == col and r['Timepoint'] == tp and r['significant']]
                
                if tp_posthoc:
                    # Find the max y value at this timepoint for positioning brackets
                    max_y = 0
                    for group_label in groups_with_data:
                        if group_label in timepoint_data[tp][col]:
                            vals = timepoint_data[tp][col][group_label]
                            if len(vals) > 0:
                                mean_val = np.mean(vals)
                                sem_val = np.std(vals) / np.sqrt(len(vals)) if len(vals) > 1 else 0
                                max_y = max(max_y, mean_val + sem_val)
                    
                    # Sort significant results by p-value to show most significant first
                    tp_posthoc_sorted = sorted(tp_posthoc, key=lambda x: x['p_value'])
                    
                    # Draw significance brackets (limit to top 5 to avoid clutter)
                    bracket_height = max_y * 0.03
                    y_base = max_y * 1.05
                    
                    for sig_idx, sig_result in enumerate(tp_posthoc_sorted[:5]):
                        stars = "***" if sig_result['p_value'] < 0.001 else "**" if sig_result['p_value'] < 0.01 else "*"
                        
                        group1 = sig_result['Group1']
                        group2 = sig_result['Group2']
                        
                        # Get x positions for both groups
                        if group1 in group_x_positions[tp] and group2 in group_x_positions[tp]:
                            x1 = group_x_positions[tp][group1]
                            x2 = group_x_positions[tp][group2]
                            
                            # Ensure x1 < x2
                            if x1 > x2:
                                x1, x2 = x2, x1
                            
                            # Calculate y position for this bracket
                            y_pos = y_base + (sig_idx * max_y * 0.08)
                            
                            # Draw the bracket
                            draw_significance_bracket(ax, x1, x2, y_pos, bracket_height, stars)
            
            ax.set_xticks(x_positions)
            ax.set_xticklabels([f'{tp} min' for tp in target_timepoints])
            ax.set_xlabel('Time from Treatment', fontsize=11)
            ax.set_ylabel(col, fontsize=11)
            ax.set_title(f'{col}', fontsize=12, fontweight='bold')
            ax.legend(loc='upper right', fontsize=8, ncol=2)
            ax.grid(True, alpha=0.3, axis='y')
            
            # Adjust y-axis to accommodate brackets
            current_ylim = ax.get_ylim()
            ax.set_ylim(current_ylim[0], current_ylim[1] * 1.4)
        
        plt.suptitle(f'{figure_title}\n(* p<0.05, ** p<0.01, *** p<0.001 - Horizontal lines connect significantly different groups)', fontsize=13, fontweight='bold')
        plt.tight_layout()
        
        filename = f'{timestamp} {filename_suffix}'
        plt.savefig(os.path.join(plots_dir, filename), dpi=150, bbox_inches='tight')
        plt.close()
        gc.collect()
        print(f"  Saved: {filename}")
    
    # ========================================================================
    # Create summary statistics comparison
    # ========================================================================
    print("\n" + "="*70)
    print("SUMMARY: Experiments Compared")
    print("="*70)
    
    for exp_name, exp_data in all_experiment_data.items():
        print(f"\n{exp_name}:")
        print(f"  Groups: {exp_data['groups']}")
        
        for group in exp_data['groups']:
            baseline = exp_data['baseline'].get(group)
            experiment = exp_data['experiment'].get(group)
            
            if baseline is not None and experiment is not None:
                baseline_filtered = get_baseline_data(baseline)
                
                for col in selected_cols:
                    if col in baseline_filtered.columns and col in experiment.columns:
                        bl_mean = baseline_filtered[col].dropna().mean()
                        exp_mean = experiment[col].dropna().mean()
                        if not np.isnan(bl_mean) and bl_mean != 0:
                            diff = ((exp_mean - bl_mean) / bl_mean) * 100
                            print(f"    {group} - {col}: Baseline={bl_mean:.2f}, Exp={exp_mean:.2f}, Diff={diff:+.1f}%")

if __name__ == "__main__":
    main()

