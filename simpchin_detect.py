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
# results=[]
for i in os.listdir(base_path):
    extension=i.split('.')[-1]
    if extension=='docx':
        text_doc=docx2txt.process(i)
        chinese_text=set(re.sub('[^%s]' % all_chinese,'',text_doc))
        if len(chinese_text)==0:
            print('{} does not contain Chinese text'.format(i))
        elif chinese_text.issubset(trad):
            print('{} has only traditional chars'.format(i))
        elif chinese_text.issubset(simp):
            print('ERROR {} is written in Simplified Chinese'.format(i))
        else:
            results=open('output_'+i.split('.')[0]+'.txt','a',encoding='utf8')
            for char in chinese_text:
                if char in simp:
                    results.write(char+'\n')
            results.close()
            #         results.append(char)
            # output=results_ser.value_counts()
            # output=pd.DataFrame({'char':output.index,'count':output.values})
            # output.to_csv(base_path+'\\output_'+i.split('.')[0]+'.txt',index=False)
            print('{} has both simplified and traditional chars'.format(i))
    else:
        print('{} is not a docx file'.format(i))