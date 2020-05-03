#!/usr/bin/env python
# coding: utf-8

# In[7]:


import urllib.request as req
import bs4
import json

# 建立 ssl連線
import ssl
ssl._create_default_https_context = ssl._create_unverified_context
# 讀取excel檔案並存檔
import pandas as pd


# # Juniper

# In[8]:


def Juniper():
    url_juniper = "https://support.juniper.net/support/eol/"
    userAgent_juniper = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.108 Safari/537.36"
    #"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.163 Safari/537.36"
    request_juniper = req.Request(url_juniper, headers = {
        "user-agent": userAgent_juniper
    })
#     ssl._create_default_https_context = ssl._create_unverified_context
    
    with req.urlopen(request_juniper) as juniper_res:
        webCode_juniper = juniper_res.read().decode("utf-8")
    root_juniper = bs4.BeautifulSoup(webCode_juniper, "lxml")
    juniper_scripts = root_juniper.find_all("script")
    juniper_eol_scripts = juniper_scripts[8].prettify()

    # 處理字串
    juniper_start_string = '{ \n\t\t\t\t\t\t"label" : "Hardware End of Life (EOL) Dates & Milestones"'
    juniper_start_index = juniper_eol_scripts.find(juniper_start_string)
    juniper_list_string = juniper_eol_scripts[juniper_start_index:]
    juniper_end_string = ']\n\t\t\t\t\t}'
    juniper_end_index = juniper_list_string.find(juniper_end_string) + len(juniper_end_string)
    juniper_list_string = juniper_list_string[:juniper_end_index]

    # 轉JSON處裡
    juniper_eol_json = json.loads(juniper_list_string)
    juniper_eol_list = juniper_eol_json["items"]

    # 處理Juniper各型號網址，取得EOL、EOS
    juniper_pre_urls = "https://support.juniper.net/"
    
    for i in juniper_eol_list:
        juniper_label_request = req.Request(juniper_pre_urls+i["url"], headers = {
            "user-agent": userAgent_juniper
        })
        with req.urlopen(juniper_label_request) as juniper_eol_res:
            juniper_eol_data = juniper_eol_res.read().decode("utf-8")
        juniper_eol_root = bs4.BeautifulSoup(juniper_eol_data, "lxml")

        # 網頁以 script布置，因此需再次處理script，並找出table內的EOL、EOS
        juniper_eol_list = juniper_eol_root.find_all("script")
        juniper_eol_table_script = juniper_eol_list[8].prettify()

        # 處理script字串(開頭index)
        juniper_eol_start_string = '<table'
        juniper_eol_start_index = juniper_eol_table_script.find(juniper_eol_start_string)
        juniper_eol_table = juniper_eol_table_script[juniper_eol_start_index:]

        # 處理script字串(結尾index)
        juniper_eol_end_string = '</table>'
        juniper_eol_end_index = juniper_eol_table.find(juniper_eol_end_string) + len(juniper_eol_end_string)
        juniper_eol_table = juniper_eol_table[:juniper_eol_end_index]

        # 使用 bs4進行處裡
        juniper_eol_table = bs4.BeautifulSoup(juniper_eol_table ,"lxml")
        juniper_eol_tr = juniper_eol_table.find_all("tr")

        # 處裡每頁的內容
        juniper_products_list = []
        product_list = []
        eol_list = []
        eos_list = []
        for tr in juniper_eol_tr:
            juniper_products = []
            juniper_eol_td = tr.find_all("td")
            for td in juniper_eol_td:
                juniper_products.append(td.string)
            juniper_products_list.append(juniper_products)

        try:
            juniper_products_list.pop(0)
            for a in range(len(juniper_products_list)):
                product_list.append(juniper_products_list[a][0])
                eol_list.append(juniper_products_list[a][1])
                eos_list.append(juniper_products_list[a][6])

        except:
            pass
        for b in range(len(product_list)):
            try:
                products = product_list[b].split(",")
                eol = eol_list[b]
                eos = eos_list[b]
            except:
                pass
            for c in range(len(products)):
                products[c] = products[c].lstrip()
                juniper_labels.append(products[c])
                juniper_eols.append(eol)
                juniper_eoss.append(eos)
#                 print(products[c])
#                 print("eol: ",eol)
#                 print("eos: ",eos)
#                 print("---------")

    # print(len(juniper_labels))
    # print(len(juniper_eols))
    # print(len(juniper_eoss))


# # Paloalto

# In[9]:


def Paloalto():
    url_pal = "https://www.paloaltonetworks.com/services/support/end-of-life-announcements/hardware-end-of-life-dates"
    userAgent_pal = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.108 Safari/537.36"
    #"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.163 Safari/537.36"
    request_pal = req.Request(url_pal, headers = {
        "user-agent" : userAgent_pal
    })
#     ssl._create_default_https_context = ssl._create_unverified_context
    
    with req.urlopen(request_pal) as res_pal:
        webCode_pal = res_pal.read().decode("utf-8")

    root_pal = bs4.BeautifulSoup(webCode_pal, "lxml")
    table_pal = root_pal.find_all("td")
    
    for i in table_pal:
        index_pal = table_pal.index(i)
        if index_pal % 5 == 0:
            label = i.text
            label = label.replace("\n", "")
            pal_labels.append(label)
        elif index_pal % 5 == 1:
            eos = i.text
            pal_eoss.append(eos)
        elif index_pal % 5 == 2:
            eol = i.text
            pal_eols.append(eol)

    # print(pal_labels)
    # print(pal_eols)
    # print(pal_eoss)


# # Cisco

# In[10]:


def Cisco():
    urls_cisco = [['Cisco 12404','https://www.cisco.com/c/en/us/products/collateral/routers/12000-series-routers/end_of_life_notice_c51-456801.html'],
                   ['Cisco 7613','https://www.cisco.com/c/en/us/products/collateral/routers/7600-series-routers/end_of_life_notice_c51-728933.html'],
                   ['Cisco ASR 1002','https://www.cisco.com/c/en/us/products/collateral/routers/asr-1000-series-aggregation-services-routers/eos-eol-notice-c51-734572.html'],
                   ['Cisco 3750X','https://www.cisco.com/c/en/us/products/collateral/switches/catalyst-3750-x-series-switches/eos-eol-notice-c51-737191.html'],
                   ['Cisco Catalyst 2960S','https://www.cisco.com/c/en/us/products/collateral/switches/catalyst-2960-series-switches/eos-eol-notice-c51-733348.html'],
                   ['Cisco Nexus 9336','https://www.cisco.com/c/en/us/products/collateral/switches/nexus-9000-series-switches/eos-eol-notice-c51-741302.html'],
                   ['Cisco 5500','https://www.cisco.com/c/en/us/products/collateral/wireless/aironet-1130-ag-series/end_of_life_notice_c51-717227.html?dtid=osscdc000283']
                  ]
    userAgent_cisco = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.108 Safari/537.36"
    
#     ssl._create_default_https_context = ssl._create_unverified_context

    for page in urls_cisco:
        cisco_labels.append(page[0])
        request_cisco = req.Request(page[1], headers = {
            "user-agent" : userAgent_cisco
        })

        with req.urlopen(request_cisco) as res_cisco:
            webCode_cisco = res_cisco.read().decode("utf-8")
        root_cisco = bs4.BeautifulSoup(webCode_cisco, "lxml")
        cisco_tables = root_cisco.find_all("td")
        cisco_table = cisco_tables[3:12]
        for i in cisco_table:
            cisco_list_index = cisco_table.index(i)
            if cisco_list_index % 9 == 2:
                eol = i.text
                eol = eol.replace(" ", "")
                cisco_eols.append(eol)
            elif cisco_list_index % 9 == 5:
                eos = i.text
                eos = eos.replace(" ", "")
                cisco_eoss.append(eos)

#     #Cisco 5500
#     cisco5500_url = "https://www.cisco.com/c/en/us/support/wireless/5500-series-wireless-controllers/tsd-products-support-series-home.html"
#     request_cisco5500 = req.Request(cisco5500_url,headers = {
#         "user-agent" : userAgent_cisco
#     })
#     with req.urlopen(request_cisco5500) as res_cisco5500:
#         webCode_cisco5500 = res_cisco5500.read().decode("utf-8")
#     root_cisco5500 = bs4.BeautifulSoup(webCode_cisco5500, "lxml")
#     cisco_labels.append("Cisco 5500")
#     cisco5500_data = root_cisco5500.find_all("dt")
#     cisco5500_eol = "None"
#     cisco5500_eos = cisco5500_data[1].text
#     cisco5500_target = "None Announced"
#     cisco5500_index = cisco5500_eos.index(cisco5500_target)
#     cisco5500_eos = cisco5500_eos[cisco5500_index:cisco5500_index+len(cisco5500_target)]
#     cisco_eols.append(cisco5500_eol)
#     cisco_eoss.append(cisco5500_eos)

    # print(cisco_labels)
    # print(cisco_eols)
    # print(cisco_eoss)


# In[11]:


# 資料儲存
juniper_labels = []
juniper_eols = []
juniper_eoss = []

pal_labels = []
pal_eols = []
pal_eoss = []

cisco_labels = []
cisco_eols = []
cisco_eoss = []

# 爬蟲
Juniper()
Paloalto()
Cisco()

#讀取excel檔案
df_item_list = pd.read_excel("file.xlsx") 
df_item_list.columns=['company', 'type'] #更改欄位名稱，方便存取

eos = []
eol = []
for row in df_item_list.itertuples():
    if row.company == 'Cisco':
        #print("###")
        #直接搜尋特定型號的EoL、EoS
        try:
            temp_index = cisco_labels.index(row.type)
            #print(row.type)
            #print(cisco_eols[temp_index])
            #print(cisco_eoss[temp_index])
            eos.append(cisco_eoss[temp_index])
            eol.append(cisco_eols[temp_index])
        #如果型號名稱有些差異，就比對網路上爬下來的資料跟目標型號的字串是否有部分符合
        except ValueError:
            for c_item in cisco_labels:
                if row.type in c_item:
                    temp_index = cisco_labels.index(c_item)
                    #print(c_item)
                    #print(cisco_eols[temp_index])
                    #print(cisco_eoss[temp_index])
                    eos.append(cisco_eoss[temp_index])
                    eol.append(cisco_eols[temp_index])

    elif row.company == 'Paloalto':
        #print("###")
        #直接搜尋特定型號的EoL、EoS
        try:
            temp_index = pal_labels.index(row.type)
            #print(row.type)
            #print(pal_eols[temp_index])
            #print(pal_eoss[temp_index])
            eos.append(pal_eoss[temp_index])
            eol.append(pal_eols[temp_index])
        #如果型號名稱有些差異，就比對網路上爬下來的資料跟目標型號的字串是否有部分符合
        except ValueError:
            for p_item in pal_labels:
                if row.type in p_item:
                    temp_index = pal_labels.index(p_item)
                    #print(p_item)
                    #print(pal_eols[temp_index])
                    #print(pal_eoss[temp_index])
                    eos.append(pal_eoss[temp_index])
                    eol.append(pal_eols[temp_index])

    elif row.company == 'Juniper':
        #print("###")
        #直接搜尋特定型號的EoL、EoS
        try:
            temp_index = juniper_labels.index(row.type)
            #print(row.type)
            #print(juniper_eols[temp_index])
            #print(juniper_eoss[temp_index])
            eos.append(juniper_eoss[temp_index])
            eol.append(juniper_eols[temp_index])
        #如果型號名稱有些差異，就比對網路上爬下來的資料跟目標型號的字串是否有部分符合
        except ValueError:
            for j_item in juniper_labels:
                if row.type in j_item:
                    #print(j_item)
                    #print(juniper_eols[temp_index])
                    #print(juniper_eoss[temp_index])
                    temp_index = juniper_labels.index(j_item)
                    eos.append(juniper_eoss[temp_index])
                    eol.append(juniper_eols[temp_index])

df_item_list['EoL'] = eol
df_item_list['EoS'] = eos

#存檔
df_item_list.to_excel('Report.xlsx', index=False)


# In[12]:


# 整合
if __name__ == '__main__':
    # 資料儲存
    juniper_labels = []
    juniper_eols = []
    juniper_eoss = []

    pal_labels = []
    pal_eols = []
    pal_eoss = []

    cisco_labels = []
    cisco_eols = []
    cisco_eoss = []
    
    # 爬蟲
    Juniper()
    Paloalto()
    Cisco()
    
    #讀取excel檔案
    df_item_list = pd.read_excel("file.xlsx") 
    df_item_list.columns=['company', 'type'] #更改欄位名稱，方便存取

    eos = []
    eol = []
    for row in df_item_list.itertuples():
        if row.company == 'Cisco':
            #print("###")
            #直接搜尋特定型號的EoL、EoS
            try:
                temp_index = cisco_labels.index(row.type)
                #print(row.type)
                #print(cisco_eols[temp_index])
                #print(cisco_eoss[temp_index])
                eos.append(cisco_eoss[temp_index])
                eol.append(cisco_eols[temp_index])
            #如果型號名稱有些差異，就比對網路上爬下來的資料跟目標型號的字串是否有部分符合
            except ValueError:
                for c_item in cisco_labels:
                    if row.type in c_item:
                        temp_index = cisco_labels.index(c_item)
                        #print(c_item)
                        #print(cisco_eols[temp_index])
                        #print(cisco_eoss[temp_index])
                        eos.append(cisco_eoss[temp_index])
                        eol.append(cisco_eols[temp_index])

        elif row.company == 'Paloalto':
            #print("###")
            #直接搜尋特定型號的EoL、EoS
            try:
                temp_index = pal_labels.index(row.type)
                #print(row.type)
                #print(pal_eols[temp_index])
                #print(pal_eoss[temp_index])
                eos.append(pal_eoss[temp_index])
                eol.append(pal_eols[temp_index])
            #如果型號名稱有些差異，就比對網路上爬下來的資料跟目標型號的字串是否有部分符合
            except ValueError:
                for p_item in pal_labels:
                    if row.type in p_item:
                        temp_index = pal_labels.index(p_item)
                        #print(p_item)
                        #print(pal_eols[temp_index])
                        #print(pal_eoss[temp_index])
                        eos.append(pal_eoss[temp_index])
                        eol.append(pal_eols[temp_index])

        elif row.company == 'Juniper':
            #print("###")
            #直接搜尋特定型號的EoL、EoS
            try:
                temp_index = juniper_labels.index(row.type)
                #print(row.type)
                #print(juniper_eols[temp_index])
                #print(juniper_eoss[temp_index])
                eos.append(juniper_eoss[temp_index])
                eol.append(juniper_eols[temp_index])
            #如果型號名稱有些差異，就比對網路上爬下來的資料跟目標型號的字串是否有部分符合
            except ValueError:
                for j_item in juniper_labels:
                    if row.type in j_item:
                        #print(j_item)
                        #print(juniper_eols[temp_index])
                        #print(juniper_eoss[temp_index])
                        temp_index = juniper_labels.index(j_item)
                        eos.append(juniper_eoss[temp_index])
                        eol.append(juniper_eols[temp_index])

    df_item_list['EoL'] = eol
    df_item_list['EoS'] = eos

    #存檔
    df_item_list.to_excel('Report.xlsx', index=False)


# In[ ]:




