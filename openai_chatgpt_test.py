#-*-coding:utf-8-*-
# Note: you need to be using OpenAI Python v0.27.0 for the code below to work
import openai
import pandas
import openpyxl
import pandas as pd
import xlrd
import argparse
import sys

## openapi, pandas, xlrd, openpyxl

parser = argparse.ArgumentParser(description="OpenAI Chat GPT Test")
parser.add_argument('--key', default='', help='open.ai에서 발급받은 인증키')
args = parser.parse_args()
openai.api_key= args.key

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
        try:
            answer = request_to_chatgpt(excel_df['request'].loc[i])
            excel_df['response'].loc[i] = answer
        except openai.error.AuthenticationError:
            print("open.ai에서 발급받은 인증키를 넣어주세요.")
            sys.exit(1)

    excel_df.to_excel(file_url, sheet_name='Sheet1', index=False)




