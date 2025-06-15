import pandas as pd

def create_readme():
    df = pd.read_excel('./source/Journals.xlsx', sheet_name="Q1-Q2", 
                       na_values="", dtype=str)
    df['Journal'] = "[" + df['Journal'] + "](" + df['Link'] + ")"
    markdown_table = df[["Journal", "Q", "APC", "Comments + Turnaround time"]].to_markdown(index=False)

    with open('./source/Canvas.md', 'r') as file:
        content = file.readlines()

    start_index = content.index("## Q1-Q2 Journals\n") + 1
    content.insert(start_index, "\n### Main Journal Table\n" + markdown_table + "\n\n")

    with open('./README.md', 'w') as file:
        file.writelines(content)

if __name__ == "__main__":
    create_readme()