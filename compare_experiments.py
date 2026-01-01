import pandas as pd
import matplotlib
matplotlib.use('Agg')  # Use non-interactive backend
import matplotlib.pyplot as plt
import numpy as np
from scipy import stats
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

