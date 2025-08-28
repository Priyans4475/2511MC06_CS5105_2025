import pandas as pd
import os


INPUT_FILE = "students.xlsx"          
BRANCH_FOLDER = "full_branch_wise"    
MIX_FOLDER = "group_branch_wise_mix" 
UNIFORM_FOLDER = "group_uniform_mix"  
FINAL_OUTPUT = "output.xlsx"         




def students_branch_wise(file_path, output_folder=BRANCH_FOLDER):
    df = pd.read_excel(file_path)

    if 'Roll' not in df.columns:
        raise ValueError("The Excel file must contain a 'Roll' column.")

    # Extract branch from roll number
    df['Branch'] = df['Roll'].astype(str).str[4:6]

    os.makedirs(output_folder, exist_ok=True)

    for branch, group in df.groupby('Branch'):
        file_name = f"{branch}.csv"
        output_path = os.path.join(output_folder, file_name)
        group[['Roll', 'Name', 'Email']].to_csv(output_path, index=False)
        print(f"[Code1] Saved {output_path}")



def students_group_mix(branch_folder=BRANCH_FOLDER,
                       output_folder=MIX_FOLDER,
                       num_groups=3):

    branch_files = sorted([f for f in os.listdir(branch_folder) if f.endswith(".csv")])
    if not branch_files:
        raise ValueError(f"No CSV files found in {branch_folder}")

    branches = {}
    for f in branch_files:
        branch = os.path.splitext(f)[0]
        df = pd.read_csv(os.path.join(branch_folder, f), dtype=str)
        df["Branch"] = branch
        branches[branch] = df.to_dict("records")

    total_students = sum(len(students) for students in branches.values())
    base, rem = divmod(total_students, num_groups)
    capacities = [base + (1 if i < rem else 0) for i in range(num_groups)]

    groups = [[] for _ in range(num_groups)]

    for g_idx in range(num_groups):
        cap = capacities[g_idx]
        while len(groups[g_idx]) < cap and any(branches.values()):
            for b in list(branches.keys()):
                if len(groups[g_idx]) >= cap:
                    break
                if branches[b]:
                    student = branches[b].pop(0)
                    groups[g_idx].append(student)

    os.makedirs(output_folder, exist_ok=True)
    for i, rows in enumerate(groups, start=1):
        gdf = pd.DataFrame(rows, columns=["Roll", "Name", "Email", "Branch"])
        gdf.to_csv(os.path.join(output_folder, f"g{i}.csv"), index=False)
        print(f"[Code2] Saved {output_folder}/g{i}.csv (total={len(gdf)})")



def students_group_uniform(branch_folder=BRANCH_FOLDER,
                           output_folder=UNIFORM_FOLDER,
                           num_groups=3):
    # Collect branch CSVs
    branch_files = [f for f in os.listdir(branch_folder) if f.endswith(".csv")]
    if not branch_files:
        raise ValueError(f"No CSV files found in {branch_folder}")

    # Load all branches into dict {branch: df}
    branches = {}
    for f in branch_files:
        branch = os.path.splitext(f)[0]
        df = pd.read_csv(os.path.join(branch_folder, f), dtype=str)
        df["Branch"] = branch
        branches[branch] = df

    # Sort branches by size (largest first)
    sorted_branches = sorted(branches.items(), key=lambda x: len(x[1]), reverse=True)

    # Compute group sizes
    total_students = sum(len(df) for df in branches.values())
    group_size = total_students // num_groups
    remainder = total_students % num_groups
    group_limits = [group_size + (1 if i < remainder else 0) for i in range(num_groups)]

    print(f"[Code3] Total students: {total_students}")
    print(f"[Code3] Target group size: {group_size} (+1 for first {remainder} groups)")
    print(f"[Code3] Group limits: {group_limits}")

    # Prepare groups
    groups = [[] for _ in range(num_groups)]
    group_idx = 0
    current_count = 0

    # Distribute students, branch by branch (largest → smallest)
    for branch, df in sorted_branches:
        for _, student in df.iterrows():
            if current_count >= group_limits[group_idx]:
                # move to next group
                group_idx += 1
                current_count = 0
                if group_idx >= num_groups:
                    raise ValueError("More students than groups capacity (logic error)")
            groups[group_idx].append(student)
            current_count += 1

    # Save each group
    os.makedirs(output_folder, exist_ok=True)
    for i, students in enumerate(groups, start=1):
        if students:
            gdf = pd.DataFrame(students)
            gdf.to_csv(os.path.join(output_folder, f"g{i}.csv"), index=False)
            print(f"[Code3] Saved {output_folder}/g{i}.csv (total={len(gdf)})")


# function
def generate_branch_stats(input_folder):
    group_files = [f for f in os.listdir(input_folder) if f.endswith(".csv")]
    if not group_files:
        raise ValueError(f"No CSV files found in {input_folder}")

    stats_list = []
    all_branches = set()

    for f in group_files:
        df = pd.read_csv(os.path.join(input_folder, f), dtype=str)
        all_branches.update(df["Branch"].unique())

    all_branches = sorted(all_branches)
    all_columns = all_branches + ["Total"]

    # preserve natural ordering of g1, g2, ..., g10, g11, etc.
    for f in sorted(group_files, key=lambda x: (x[0], int(x[1:].split('.')[0]) if x[1:].split('.')[0].isdigit() else 999)):
        group_name = os.path.splitext(f)[0]
        df = pd.read_csv(os.path.join(input_folder, f), dtype=str)
        counts = df["Branch"].value_counts().to_dict()
        row = [counts.get(branch, 0) for branch in all_branches]
        row.append(len(df))
        stats_list.append([group_name] + row)

    return stats_list, all_columns


def Generate_output(mix_folder=MIX_FOLDER, uniform_folder=UNIFORM_FOLDER, output_excel=FINAL_OUTPUT):
    mix_stats, mix_columns = generate_branch_stats(mix_folder)
    uniform_stats, uniform_columns = generate_branch_stats(uniform_folder)

    all_columns = sorted(set(mix_columns + uniform_columns))
    final_columns = [""] + all_columns 

    mix_header = [["Mix"] + all_columns]
    uniform_header = [["Uniform"] + all_columns]

    mix_rows = []
    for row in mix_stats:
        row_dict = dict(zip(["Group"] + mix_columns, row))
        mix_rows.append([row_dict["Group"]] + [row_dict.get(col, 0) for col in all_columns])

    uniform_rows = []
    for row in uniform_stats:
        row_dict = dict(zip(["Group"] + uniform_columns, row))
        uniform_rows.append([row_dict["Group"]] + [row_dict.get(col, 0) for col in all_columns])

    blank_rows = [[""] * len(final_columns)] * 2

    final_data = mix_header + mix_rows + blank_rows + uniform_header + uniform_rows
    final_df = pd.DataFrame(final_data)

    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        final_df.to_excel(writer, sheet_name="Stats", index=False, header=False)

    print(f"[Code4] Stats saved to {output_excel} (sheet=Stats)")



if __name__ == "__main__":
    

    # ask user input
    num_groups = int(input("Enter number of groups: "))

    students_branch_wise(INPUT_FILE, BRANCH_FOLDER)
    students_group_mix(BRANCH_FOLDER, MIX_FOLDER, num_groups)
    students_group_uniform(BRANCH_FOLDER, UNIFORM_FOLDER, num_groups)
    Generate_output(MIX_FOLDER, UNIFORM_FOLDER, FINAL_OUTPUT)

    print("All folders/files generated.")
