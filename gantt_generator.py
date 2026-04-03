import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from datetime import datetime
from rag_engine import run_rag_engine

RAG_COLORS = {
    'Red': '#FF4444',
    'Amber': '#FFA500',
    'Green': '#44BB44',
    'Grey': '#AAAAAA'
}

def build_gantt(df, project_name, output_path):
    proj_df = df[df['Project_Name'] == project_name].copy()
    proj_df = proj_df.dropna(subset=['Start_Date', 'Planned_End_Date'])
    proj_df = proj_df.sort_values('Start_Date', ascending=False).reset_index(drop=True)

    today = pd.Timestamp(datetime.today().date())
    project_start = proj_df['Start_Date'].min()

    fig, ax = plt.subplots(figsize=(16, max(8, len(proj_df) * 0.5)))
    fig.patch.set_facecolor('#F8F9FA')
    ax.set_facecolor('#F8F9FA')

    for i, row in proj_df.iterrows():
        start_offset = (row['Start_Date'] - project_start).days
        duration = (row['Planned_End_Date'] - row['Start_Date']).days
        color = RAG_COLORS.get(row['RAG'], '#AAAAAA')
        lw = 3 if row['Is_Critical'] else 1

     
        ax.barh(i, duration, left=start_offset, height=0.5,
                color=color, alpha=0.85,
                edgecolor='black' if row['Is_Critical'] else '#666',
                linewidth=lw)

        
        progress_width = duration * row['Pct_Complete'] / 100
        ax.barh(i, progress_width, left=start_offset, height=0.5,
                color=color, alpha=1.0, edgecolor='none')

        
        label = f"{'★ ' if row['Is_Critical'] else ''}{row['Task_ID']}: {row['Task_Name'][:35]}"
        ax.text(-2, i, label, va='center', ha='right', fontsize=7.5, color='#222')

        
        ax.text(start_offset + duration + 0.5, i,
                f"{int(row['Pct_Complete'])}%", va='center', fontsize=7, color='#555')

    
    today_offset = (today - project_start).days
    ax.axvline(today_offset, color='#1F3864', linewidth=2, linestyle='--', label='Today')
    ax.text(today_offset + 0.3, len(proj_df) - 0.5, 'TODAY',
            color='#1F3864', fontsize=8, fontweight='bold')

    
    total_days = (proj_df['Planned_End_Date'].max() - project_start).days + 5
    tick_positions = range(0, total_days, 7)
    tick_labels = [(project_start + pd.Timedelta(days=d)).strftime('%d %b') for d in tick_positions]
    ax.set_xticks(list(tick_positions))
    ax.set_xticklabels(tick_labels, rotation=45, fontsize=8)
    ax.set_yticks([])
    ax.set_xlim(-2, total_days)

    
    legend_handles = [
        mpatches.Patch(color=RAG_COLORS['Green'], label='Green — On Track'),
        mpatches.Patch(color=RAG_COLORS['Amber'], label='Amber — At Risk'),
        mpatches.Patch(color=RAG_COLORS['Red'], label='Red — Overdue/Critical'),
        plt.Line2D([0], [0], color='#1F3864', linewidth=2, linestyle='--', label='Today'),
    ]
    ax.legend(handles=legend_handles, loc='lower right', fontsize=8)

    ax.set_title(f"Gantt Chart — {project_name}\nlemu_tajin | Generated {today.strftime('%d %b %Y')}",
                 fontsize=12, fontweight='bold', pad=12)
    ax.grid(axis='x', alpha=0.3, linestyle=':')

    plt.tight_layout()
    plt.savefig(output_path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"Gantt saved: {output_path}")

def generate_all_gantts(tasks_filepath):
    df, _, _ = run_rag_engine(tasks_filepath)
    for project in df['Project_Name'].unique():
        safe_name = project.replace(' ', '_').replace('/', '_')
        build_gantt(df, project, f"outputs/gantt_{safe_name}.png")

if __name__ == '__main__':
    generate_all_gantts('Tasks.xlsx')