#-*-coding:utf-8-*-
# Note: you need to be using OpenAI Python v0.27.0 for the code below to work
import openai
import pandas
import openpyxl
import pandas as pd
import xlrd

## openapi, pandas, xlrd, openpyxl
openai.api_key="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"

def request_to_chatgpt(req_string):

    response = openai.ChatCompletion.create(
      model="gpt-3.5-turbo-0301",
      messages=[
            {"role": "user", "content": req_string},
        ]
    )

    #print(response['usage']['total_tokens'])
    return response['choices'][0]['message']['content']

def read_excel_file(file_url):
    file_name = file_url

    df = pd.read_excel(file_name, engine='openpyxl', sheet_name=['Sheet1'])
    request_df = df['Sheet1']

    return request_df


if __name__ == "__main__":
    file_url = "request.xlsx"
    excel_df = read_excel_file(file_url)

    for i in range(excel_df[['request']].size):
        print(excel_df['request'].loc[i])
        answer = request_to_chatgpt(excel_df['request'].loc[i])
        excel_df['response'].loc[i] = answer

    excel_df.to_excel(file_url, sheet_name='Sheet1', index=False)




