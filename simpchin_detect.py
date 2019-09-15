#%%
from zhon import cedict
import re
import docx2txt
import os
from os.path import abspath,join
#%%
base_path=os.path.dirname(abspath('__file__'))
#%%
trad=set(list(cedict.traditional))
simp=set(list(cedict.simplified))
both=set([i for i in trad if i in simp])
all_chinese=cedict.all
#%%
def report(msg,filename,result):
    msg_head='*'*20+'\n'+'Result for {}:'.format(filename)+'\n'
    result_msg='RESULT :: '+result+'\n'
    msg_body=msg+'\n'+'-'*20+'\n'
    return msg_head+result_msg+msg_body
#%%
msg_list=[]
result=open('script_result.txt','a',encoding='utf8')
for i in os.listdir(base_path):
    extension=i.split('.')[-1]
    if extension==i:
        pass
    elif extension in ['py','git','spec','exe','txt','md','gitattributes']:
        pass
    elif extension=='docx':
        text_doc=docx2txt.process(i)
        chinese_text=set(re.sub('[^%s]' % all_chinese,'',text_doc))
        if len(chinese_text)==0:
            msg='{} does not contain Chinese text'.format(i)
            msg_list.append(report(msg,i,'IGNORE FILE'))
        elif chinese_text.issubset(trad):
            msg='{} is written in Traditional Chinese.\nConfirm that the service is E to TC. Otherwise, it is a serious error.'.format(i)
            msg_list.append(report(msg,i,'TRADITIONAL CHINESE'))
        elif chinese_text.issubset(simp):
            msg='{} is written in Simplified Chinese.\nConfirm that the service is E to SC. Otherwise, it is a serious error.'.format(i)
            msg_list.append(report(msg,i,'SIMPLIFIED CHINESE'))
        else:
            output=open('output_'+i.split('.')[0]+'.txt','a',encoding='utf8')
            for char in chinese_text:
                if char in simp:
                    output.write(char+'\n')
            output.close()
            output_name='output_'+i.split('.')[0]+'.txt'
            msg_strt='{} has both Simplified and Traditional characters'.format(i)+'\n'+'Check service name and fix characters of other language.'+'\n'
            msg_end=' has been generated. It is a list of Simplified Characters to be fixed.'
            full_msg=msg_strt+output_name+msg_end
            msg_list.append(report(full_msg,i,'ERROR'))
    else:
        msg='{} is not a docx file'.format(i)
        msg_list.append(report(msg,i,'IGNORE FILE'))
for i in msg_list:
    result.write(i)
result.close()