import pandas as pd
from pathlib import Path

def main():
    df = pd.read_excel(Path(r"C:\Users\lionm\Downloads\2024_RMA_SPR_xFusion발송.xlsx"), index_col='요청일')
    df = df.reset_index()
    df['요청일'] = df['요청일'].dt.strftime('%Y-%m-%d')
    #print(len(df))
    #print(df)

    df_5288V7 = df[df['Model'] == '5288 V7']
    #df_5288V7.to_excel(Path(r"C:\Users\lionm\Downloads\2024_RMA_SPR_xFusion발송_5288V7.xlsx"), index=False)
    #print(df_5288V7)
    df_5288V7_count = len(df_5288V7)
    print(len(df_5288V7))

    df_2288HV7 = df[df['Model'] == '2288H V7']
    #df_2288HV7.to_excel(Path(r"C:\Users\lionm\Downloads\2024_RMA_SPR_xFusion발송_2288HV7.xlsx"), index=False)
    df_2288HV7_count = len(df_2288HV7)
    print(len(df_2288HV7))

    df_1288HV7 = df[df['Model'] == '1288H V7']
    #df_1288HV7.to_excel(Path(r"C:\Users\lionm\Downloads\2024_RMA_SPR_xFusion발송_1288HV7.xlsx"), index=False)
    df_1288HV7_count = len(df_1288HV7)
    print(len(df_1288HV7))

    df_1288HV7_LFF = df[df['Model'] == '1288H V7(LFF)']
    #df_1288HV7_LFF.to_excel(Path(r"C:\Users\lionm\Downloads\2024_RMA_SPR_xFusion발송_1288HV7_LFF.xlsx"), index=False)
    df_1288VH7_LFF_count = len(df_1288HV7_LFF)
    print(len(df_1288HV7_LFF))

    total_count = df_5288V7_count + df_2288HV7_count + df_1288HV7_count + df_1288VH7_LFF_count
    print("Total Count:", total_count)

if __name__ == "__main__":
    main()