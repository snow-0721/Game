import requests
import json

# 定义必要的参数
tenant_id = 'YOUR_TENANT_ID'
client_id = 'YOUR_CLIENT_ID'
client_secret = 'YOUR_CLIENT_SECRET'
user_email = 'USER_EMAIL'

# 获取访问令牌
def get_access_token():
    url = 'https://login.microsoftonline.com/{0}/oauth2/token'.format(tenant_id)
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    data = {
        'grant_type': 'client_credentials',
        'client_id': client_id,
        'client_secret': client_secret,
        'resource': 'https://graph.microsoft.com/'
    }
    response = requests.post(url, headers=headers, data=data)
    access_token = response.json().get('access_token')
    return access_token

# 获取用户日历
def get_user_calendar(access_token):
    url = 'https://graph.microsoft.com/v1.0/users/{0}/calendar/events'.format(user_email)
    headers = {
        'Authorization': 'Bearer {0}'.format(access_token),
        'Content-Type': 'application/json'
    }
    response = requests.get(url, headers=headers)
    events = response.json().get('value', [])
    return events

# 主程序
def main():
    access_token = get_access_token()
    events = get_user_calendar(access_token)
    for event in events:
        print('主题：', event.get('subject'))
        print('开始时间：', event.get('start').get('dateTime'))
        print('结束时间：', event.get('end').get('dateTime'))
        print('---')

if __name__ == '__main__':
    main()
