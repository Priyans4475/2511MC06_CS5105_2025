
import streamlit as st
import pandas as pd
import os
import zipfile
import io

# ----------------- Your Functions -----------------
def students_branch_wise(file_path, output_folder):
    df = pd.read_excel(file_path)

    if 'Roll' not in df.columns:
        st.error("The Excel file must contain a 'Roll' column.")
        return

    df['Branch'] = df['Roll'].astype(str).str[4:6]
    os.makedirs(output_folder, exist_ok=True)

    for branch, group in df.groupby('Branch'):
        file_name = f"{branch}.csv"
        output_path = os.path.join(output_folder, file_name)
        group[['Roll', 'Name', 'Email']].to_csv(output_path, index=False)


def students_group_mix(branch_folder, output_folder, num_groups):
    branch_files = sorted([f for f in os.listdir(branch_folder) if f.endswith(".csv")])
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


def students_group_uniform(branch_folder, output_folder, num_groups):
    branch_files = [f for f in os.listdir(branch_folder) if f.endswith(".csv")]
    branches = {}
    for f in branch_files:
        branch = os.path.splitext(f)[0]
        df = pd.read_csv(os.path.join(branch_folder, f), dtype=str)
        df["Branch"] = branch
        branches[branch] = df

    sorted_branches = sorted(branches.items(), key=lambda x: len(x[1]), reverse=True)
    total_students = sum(len(df) for df in branches.values())
    group_size = total_students // num_groups
    remainder = total_students % num_groups
    group_limits = [group_size + (1 if i < remainder else 0) for i in range(num_groups)]

    groups = [[] for _ in range(num_groups)]
    group_idx, current_count = 0, 0

    for branch, df in sorted_branches:
        for _, student in df.iterrows():
            if current_count >= group_limits[group_idx]:
                group_idx += 1
                current_count = 0
            groups[group_idx].append(student)
            current_count += 1

    os.makedirs(output_folder, exist_ok=True)
    for i, students in enumerate(groups, start=1):
        if students:
            gdf = pd.DataFrame(students)
            gdf.to_csv(os.path.join(output_folder, f"g{i}.csv"), index=False)


def generate_branch_stats(input_folder):
    group_files = [f for f in os.listdir(input_folder) if f.endswith(".csv")]
    stats_list, all_branches = [], set()

    for f in group_files:
        df = pd.read_csv(os.path.join(input_folder, f), dtype=str)
        all_branches.update(df["Branch"].unique())

    all_branches = sorted(all_branches)
    all_columns = all_branches + ["Total"]

    for f in sorted(group_files):
        group_name = os.path.splitext(f)[0]
        df = pd.read_csv(os.path.join(input_folder, f), dtype=str)
        counts = df["Branch"].value_counts().to_dict()
        row = [counts.get(branch, 0) for branch in all_branches]
        row.append(len(df))
        stats_list.append([group_name] + row)

    return stats_list, all_columns


def Generate_output(mix_folder, uniform_folder, output_excel):
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


# ----------------- Streamlit App -----------------
st.title("📊 Student Grouping App")

uploaded_file = st.file_uploader("Upload Excel file (must contain 'Roll' column)", type=["xlsx"])
num_groups = st.number_input("Enter number of groups:", min_value=2, max_value=20, value=3, step=1)

if uploaded_file and st.button("Generate Groups"):
    # Create base output folder
    base_folder = "student_groups_output"
    os.makedirs(base_folder, exist_ok=True)

    # Save uploaded file
    input_path = os.path.join(base_folder, "students.xlsx")
    with open(input_path, "wb") as f:
        f.write(uploaded_file.read())

    # Define output folders
    branch_folder = os.path.join(base_folder, "full_branch_wise")
    mix_folder = os.path.join(base_folder, "group_branch_wise_mix")
    uniform_folder = os.path.join(base_folder, "group_uniform_mix")
    output_excel = os.path.join(base_folder, "output.xlsx")

    # Run processing
    students_branch_wise(input_path, branch_folder)
    students_group_mix(branch_folder, mix_folder, num_groups)
    students_group_uniform(branch_folder, uniform_folder, num_groups)
    Generate_output(mix_folder, uniform_folder, output_excel)


    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        for folder in [branch_folder, mix_folder, uniform_folder]:
            for root, _, files in os.walk(folder):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, base_folder)
    zip_buffer.seek(0)


    st.download_button(
        "⬇️ Download All Groups (ZIP)",
        zip_buffer,
        file_name="student_groups.zip",
        mime="application/zip"
    )

    with open(output_excel, "rb") as f:
        st.download_button(
            "⬇️ Download Final Output (Excel)",
            f,
            file_name="output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.success(f"✅ Groups generated! Download ZIP for all groups or Excel for stats.")
