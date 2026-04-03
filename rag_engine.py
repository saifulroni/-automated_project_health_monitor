import pandas as pd
from datetime import datetime, timedelta
import networkx as nx

def load_tasks(filepath):
    df = pd.read_excel(filepath, sheet_name='Tasks')

#format dates
    date_cols = ['Start_Date', 'Planned_End_Date', 'Actual_Completion_Date']
    for col in date_cols:
        df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce')
    df['Pct_Complete'] = pd.to_numeric(df['Pct_Complete'], errors='coerce').fillna(0)
    return df

#assign rag status
def assign_task_rag(row, today):
    status = row['Status']
    end = row['Planned_End_Date']
    actual = row['Actual_Completion_Date']
    pct = row['Pct_Complete']

    if status == 'Cancelled':
        return 'Grey'

    if status == 'Completed':
        if pd.notna(actual) and actual <= end:
            return 'Green'
        elif pd.notna(actual) and (actual - end).days <= 3:
            return 'Amber'
        else:
            return 'Red'

    if pd.notna(end):
        days_remaining = (end - today).days
        time_elapsed_pct = (today - row['Start_Date']).days / max((end - row['Start_Date']).days, 1) * 100

#rag conditions
        if today > end:  # Overdue
            return 'Red'
        if pct < 30 and (100 - time_elapsed_pct) < 20:
            return 'Red'

        if days_remaining <= 5 and pct < 70:
            return 'Amber'
        if status == 'Blocked':
            return 'Amber'  

    return 'Green'

#dependencies check
def get_critical_path(df):
    G = nx.DiGraph()
    for _, row in df.iterrows():
        G.add_node(row['Task_ID'], duration=(
            (row['Planned_End_Date'] - row['Start_Date']).days
            if pd.notna(row['Planned_End_Date']) and pd.notna(row['Start_Date']) else 1
        ))
    for _, row in df.iterrows():
        if pd.notna(row['Dependencies']) and str(row['Dependencies']).strip():
            deps = [d.strip() for d in str(row['Dependencies']).split(',')]
            for dep in deps:
                if dep in G.nodes:
                    G.add_edge(dep, row['Task_ID'])
    try:
# critical path
        critical_path = nx.dag_longest_path(G, weight='duration')
    except Exception:
        critical_path = []
    return critical_path

def assign_project_rag(project_df, critical_path_ids):
    task_rags = project_df['RAG'].tolist()
    critical_tasks = project_df[project_df['Task_ID'].isin(critical_path_ids)]

    if 'Red' in critical_tasks['RAG'].values:
        return 'Red'
    if 'Red' in task_rags:
        return 'Amber'
    amber_pct = task_rags.count('Amber') / len(task_rags)
    if amber_pct >= 0.2:
        return 'Amber'
    return 'Green'

#overdue comparison

def run_rag_engine(filepath):
    today = pd.Timestamp(datetime.today().date())
    df = load_tasks(filepath)
    df['RAG'] = df.apply(lambda row: assign_task_rag(row, today), axis=1)

    critical_path = get_critical_path(df)
    df['Is_Critical'] = df['Task_ID'].isin(critical_path)

#Upgrade blocked critical path tasks to Red
    df.loc[(df['Is_Critical']) & (df['Status'] == 'Blocked'), 'RAG'] = 'Red'

    project_rags = {}
    for project in df['Project_Name'].unique():
        proj_df = df[df['Project_Name'] == project]
        project_rags[project] = assign_project_rag(proj_df, critical_path)

    return df, project_rags, critical_path

#if __name__ == '__main__':
    df, proj_rags, cp = run_rag_engine('tasks.xlsx')
 #   print(df[['Task_ID','Task_Name','RAG','Is_Critical']].to_string())
  #  print("\nProject RAGs:", proj_rags)
  #  print("\nCritical Path:", cp)