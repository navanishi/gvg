import pandas as pd
import requests

# GitHubのリポジトリからファイルをダウンロード
url = 'https://github.com/navanishi/gvg/raw/main/togvg.xlsx'
file_path = 'togvg.xlsx'

try:
    response = requests.get(url)
    response.raise_for_status()  # HTTPエラーチェック
    with open(file_path, 'wb') as f:
        f.write(response.content)

    # データを読み込む
    atend_sheet = 'atend'
    memlist_sheet = 'memlist'
    party_sheet = 'party'

    atend_df = pd.read_excel(file_path, sheet_name=atend_sheet, header=None, engine='openpyxl')
    memlist_df = pd.read_excel(file_path, sheet_name=memlist_sheet, engine='openpyxl')

    # 出席メンバーの名前リスト
    atend_names = atend_df[0].tolist()  # atendシートの最初の列が名前のリスト

    # パーティ数を計算
    party_count = (len(atend_names) + 7) // 6

    # パーティシートの初期化
    party_df = pd.DataFrame(columns=['A', 'B', 'C', 'D', 'E', 'F', 'G'])
    party_df['A'] = ['party' + str(i + 1) for i in range(party_count)]

    # 配置するメンバーの管理
    assigned_leaders = set()
    assigned_members = set()

    # カテゴリ6かつリーダー（L）のメンバーを抽出
    leaders_6 = memlist_df[(memlist_df.iloc[:, 3] == 6) & (memlist_df.iloc[:, 5] == 'L')].iloc[:, 0].tolist()
    leader_count = 0
    for i in range(party_count):
        if leader_count < len(leaders_6):
            party_df.at[i, 'B'] = leaders_6[leader_count]
            assigned_leaders.add(leaders_6[leader_count])
            assigned_members.add(leaders_6[leader_count])
            print(f"Assigned category 6 leader: {leaders_6[leader_count]} to party {i + 1}")
            leader_count += 1

    # atendシートに名前があるリーダーを抽出
    other_leaders = memlist_df[(memlist_df.iloc[:, 5] == 'L') & (memlist_df.iloc[:, 0].isin(atend_names)) & (memlist_df.iloc[:, 3] != 6)].iloc[:, 0].tolist()
    for i in range(party_count):
        if pd.isna(party_df.at[i, 'B']) and leader_count < len(other_leaders):
            if other_leaders[leader_count] not in assigned_leaders:  # 他のリーダーとしては一度だけ割り当てる
                party_df.at[i, 'B'] = other_leaders[leader_count]
                assigned_leaders.add(other_leaders[leader_count])
                assigned_members.add(other_leaders[leader_count])
                print(f"Assigned other leader: {other_leaders[leader_count]} to party {i + 1}")
                leader_count += 1

    # カテゴリ10のメンバーをカテゴリ6のリーダーのパーティに配置
    members_10 = memlist_df[memlist_df.iloc[:, 3] == 10].iloc[:, 0].tolist()

    print(f"Members category 10: {members_10}")

    for member_10 in members_10:
        if member_10 not in assigned_members:
            assigned = False
            for i in range(party_count):
                if party_df.at[i, 'B'] in leaders_6:
                    for j in range(2, 7):
                        if pd.isna(party_df.at[i, party_df.columns[j]]):
                            party_df.at[i, party_df.columns[j]] = member_10
                            assigned_members.add(member_10)
                            print(f"Assigned category 10 member: {member_10} to party {i + 1}")
                            assigned = True
                            break
                    if assigned:
                        break

    # 残っているリーダーのリストから既に割り当てられたリーダーを削除
    other_leaders = [leader for leader in other_leaders if leader not in assigned_leaders]

    # カテゴリ5かつリーダー（L）のメンバーを抽出
    leaders_5 = memlist_df[(memlist_df.iloc[:, 3] == 5) & (memlist_df.iloc[:, 5] == 'L')].iloc[:, 0].tolist()
    leader_count_5 = 0
    for i in range(party_count):
        if pd.isna(party_df.at[i, 'B']) and leader_count_5 < len(leaders_5):
            if leaders_5[leader_count_5] not in assigned_leaders:  # カテゴリ5のリーダーとして一度だけ割り当てる
                party_df.at[i, 'B'] = leaders_5[leader_count_5]
                assigned_leaders.add(leaders_5[leader_count_5])
                assigned_members.add(leaders_5[leader_count_5])
                print(f"Assigned category 5 leader: {leaders_5[leader_count_5]} to party {i + 1}")
                leader_count_5 += 1

    # カテゴリ4のメンバーをカテゴリ5のリーダーのパーティに配置
    members_4 = memlist_df[(memlist_df.iloc[:, 3] == 4) & (memlist_df.iloc[:, 0].isin(atend_names))].iloc[:, 0].tolist()

    print(f"Members category 4: {members_4}")

    for member_4 in members_4:
        if member_4 not in assigned_members:
            assigned = False
            for i in range(party_count):
                if party_df.at[i, 'B'] in leaders_5:
                    for j in range(2, 7):
                        if pd.isna(party_df.at[i, party_df.columns[j]]):
                            party_df.at[i, party_df.columns[j]] = member_4
                            assigned_members.add(member_4)
                            print(f"Assigned category 4 member: {member_4} to party {i + 1}")
                            assigned = True
                            break
                    if assigned:
                        break

    # 残りのカテゴリ10のメンバーを同じパーティに配置
    remaining_members_10 = [m for m in members_10 if m not in assigned_members]

    print(f"Remaining members category 10: {remaining_members_10}")

    for member_10 in remaining_members_10:
        for i in range(party_count):
            if party_df.iloc[i, 1:7].isna().sum() > 0:
                for j in range(1, 7):
                    if pd.isna(party_df.at[i, party_df.columns[j]]):
                        party_df.at[i, party_df.columns[j]] = member_10
                        assigned_members.add(member_10)
                        print(f"Assigned remaining category 10 member: {member_10} to party {i + 1}")
                        break
                if member_10 in assigned_members:
                    break

    # 残りのメンバーを配置
    remaining_members = [name for name in atend_names if name not in assigned_members]
    print(f"Remaining members: {remaining_members}")

    for member in remaining_members:
        for i in range(party_count):
            if party_df.iloc[i, 1:7].isna().sum() > 0:
                for j in range(1, 7):
                    if pd.isna(party_df.at[i, party_df.columns[j]]):
                        party_df.at[i, party_df.columns[j]] = member
                        assigned_members.add(member)
                        print(f"Assigned remaining member: {member} to party {i + 1}")
                        break
                if member in assigned_members:
                    break

    # 最終状態のパーティシートを表示
    print("Final party_df:")
    print(party_df)

    # 結果をExcelファイルに保存
    with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='replace', engine='openpyxl') as writer:
        party_df.to_excel(writer, sheet_name=party_sheet, index=False)

except Exception as e:
    print(f"Excel 読み込みエラー: {e}")
