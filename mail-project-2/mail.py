import win32com.client
import pandas as pd

def match_team_leader(team_leader_name):
    match team_leader_name:
        case '张三': return '销售分成确认template_1.msg'
        case '李四': return '销售分成确认template_2.msg'
        case '王五': return '销售分成确认template_3.msg'
        case '赵六': return '销售分成确认template_4.msg'
        case _: raise Exception(f'Error: 无{team_leader_name}邮件模板')

def generate_mail(outlook, in_dir, out_dir, template, name, address):
    template_path = f'{in_dir}{template}'
    mail = outlook.OpenSharedItem(template_path)
    mail.To = address
    mail.Subject = f'销售分成确认--{name}'
    attachment = f'{in_dir}销售分成确认_按销售拆分\\销售分成_{name}.xlsx'
    mail.Attachments.Add(attachment)
    out_filename = f'{out_dir}销售分成确认_{name}.msg'
    mail.SaveAs(out_filename)


in_dir = 'G:\\_Private\\EQUITY_DERIVATIVES\\Intern\\Dir1\\'
out_dir = 'C:\\Users\\LWY\\Downloads\\'

filename_contacts = '销售列表 - 1.xlsx'

df_raw = pd.read_excel(f'{in_dir}{filename_contacts}')

df = df_raw[['姓名', 'To']]
df['template'] = df_raw['团队长'].apply(match_team_leader)

outlook = win32com.client.Dispatch("outlook.application").GetNamespace("MAPI")

for _, row in df.iterrows():
    name = row['姓名']
    address = row['To']
    template = row['template']
    generate_mail(outlook, in_dir, out_dir, template, name, address)


# mail = ol.CreateItem(0)

# mail.Subject = 'Testing Mail'
# mail.To = 'intern14@xxxx.com'
# mail.CC = 'intern14@xxxx.com'

# mail.HTMLBody = r"""<p style="font-family: Calibri; font-size: 11pt">
#     各位销售老师，<br><br>
#     为协助部门XX系统流程优化，烦请<font color="red"><b>确认附件中的各类信息</b></font>，特别注意<font color="red"><b>“D列：新Sales数据”中的分成信息</b></font>是否准确；<br><br>
#     此后，销售分成等数据将一并上传XX系统，届时再修改需要从DAS发起修改审批等相关流程，更为稳健但<font color="red"><b>耗时会相对延长</b></font>，故辛苦大家尽量核对现有数据无误以提高效率；<br><br>
#     若有任何缺漏或不一致信息，请与相关团队确认后联系产品组进行修改，谢谢！<br><br>
#     Best,<br><br>
#     </p>
# """
