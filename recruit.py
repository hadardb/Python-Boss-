import requests
from pyquery import PyQuery as pq
import re
import xlwt
import time
import webbrowser

#设置头部防止被黑
headers ={
    'User-Agent':'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.110 Safari/537.36'
    }

# 书写xls
def make_excel(jobExc,info):
    if jobExcList.count(jobExc) == 0:
        jobExcList.append(jobExc)
        table = file.add_sheet(jobExc, cell_overwrite_ok=True)
        table.col(0).width = (30 * 200)     #岗位名称
        table.col(1).width = (30 * 200)     #地址
        table.col(2).width = (30 * 100)     #学历
        table.col(3).width = (30 * 200)     #单位
        table.col(4).width = (30 * 100)     #工资
        table.col(5).width = (30 * 100)     #平均工资
        table.col(6).width = (30 * 100)  # 地址
        index[jobExc]=0
        AverageList[jobExc]=0
        i=0
        while(i<7):
            table.write(0,i,title[i] )
            i=i+1
    else:
        table = file.get_sheet(jobExc)
    index[jobExc]=index[jobExc]+1
    AverageList[jobExc]=AverageList[jobExc]+(int(re.sub("\D", "",info[4].split('-')[0]))+int(re.sub("\D", "",info[4].split('-')[1])))
    i = 0
    while (i < len(info)):
        table.write(index[jobExc],i,info[i])
        i = i + 1
#写平均工资
def make_average():
    style = xlwt.XFStyle()  # 格式信息
    font = xlwt.Font()  # 字体基本设置
    font.name = u'微软雅黑'
    font.color = 'red'
    font.height = 220  # 字体大小，220就是11号字体，大概就是11*20得来的吧
    style.font = font
    alignment = xlwt.Alignment()  # 设置字体在单元格的位置
    alignment.horz = xlwt.Alignment.HORZ_CENTER  # 水平方向
    alignment.vert = xlwt.Alignment.VERT_CENTER  # 竖直方向
    style.alignment = alignment
    for key in AverageList:
        table = file.get_sheet(key)
        table.write_merge(1, index[key], 6, 6, round(AverageList[key] / index[key] / 2, 2), style)
# 写echarts
def make_echarts(name,index):
    value=""
    for key in index:
        value=value+"{value:%d, name:'%s'}," % (index[key],key)
    # 命名生成的html
    GEN_HTML = "%s.html" %(name)
    # 打开文件，准备写入
    f = open(GEN_HTML, 'w', encoding='utf8')
    message = """
    <!DOCTYPE html>
    <html>
    	<head>
    		<meta charset="UTF-8">
    		<title></title>
    		<script src="https://cdn.bootcss.com/echarts/4.2.0-rc.2/echarts-en.common.min.js"></script>
    	</head>
    	<body>
    		 <div id="main" style="width: 800px;height:600px;margin: 100px auto;"></div>
        <script type="text/javascript">
            var myChart = echarts.init(document.getElementById('main'));
            var option = {
        backgroundColor: '#2c343c',
        title: {
            text: '%s',
            left: 'center',
            top: 20,
            textStyle: {
                color: '#ccc'
            }
        },
        tooltip : {
            trigger: 'item',
            formatter: "{a} <br/>{b} : {c} ({d}%%)"
        },
        visualMap: {
            show: false,
            min: 80,
            max: 600,
            inRange: {
                colorLightness: [0, 1]
            }
        },
        series : [
            {
                name:'工作经验',
                type:'pie',
                radius : '55%%',
                center: ['50%%', '50%%'],
                data:[
                    %s
                ].sort(function (a, b) { return a.value - b.value; }),
                roseType: 'radius',
                label: {
                    normal: {
                        textStyle: {
                            color: 'rgba(255, 255, 255, 0.3)'
                        }
                    }
                },
                labelLine: {
                    normal: {
                        lineStyle: {
                            color: 'rgba(255, 255, 255, 0.3)'
                        },
                        smooth: 0.2,
                        length: 10,
                        length2: 20
                    }
                },
                itemStyle: {
                    normal: {
                        color: '#c23531',
                        shadowBlur: 200,
                        shadowColor: 'rgba(0, 0, 0, 0.5)'
                    }
                },
                animationType: 'scale',
                animationEasing: 'elasticOut',
                animationDelay: function (idx) {
                    return Math.random() * 200;
                }
            }
        ]
    };
            myChart.setOption(option);
        </script>
    	</body>
    </html>
    """ % (name+'招聘人数统计', value)
    # 写入文件
    f.write(message)
    # 关闭文件
    f.close()
    # 运行完自动在网页中显示
    webbrowser.open(GEN_HTML, new=1)

# 获取地址与信息
def set_response(name,index):
    response = requests.get("https://www.zhipin.com/c101190400/?query=%s&page=%d&ka=page-%d" % (name,index,index), headers=headers)
    doc = pq(response.text)
    doc = pq(doc('.job-list ul li')).items()
    for li in doc:
        workName=pq(li('.job-title'))
        place=pq(li('.info-primary p '))
        wages=pq(li('.red'))
        company=pq(li('.company-text .name'))
        url=pq(li('.name a'))
        matchObj = re.match('^<p>(.*)<em class="vline"/>(.*)<em class="vline"/>(.*)</p>',str(place))
        make_excel(matchObj.group(2),[workName.text(),matchObj.group(1),matchObj.group(3),company.text(),wages.text(),'https://www.zhipin.com/%s' % (url.attr('href'))])
        # print('岗位：'+workName.text())
        # print('地址：'+matchObj.group(1))
        # print('工作经验：'+matchObj.group(2))
        # print('学历：'+matchObj.group(3))
        # print((int(re.sub("\D", "",wages.text().split('-')[0]))+int(re.sub("\D", "",wages.text().split('-')[1])))/2)
        # print('工资：'+wages.text())
        # print('单位：'+company.text())
if __name__ == '__main__':
    print('请输入需要查找的岗位名称：')
    name = input()
    ecahrtsName='boss直聘江苏%s%s' % (name,time.strftime("%Y%m%d"))
    xlsName='boss直聘江苏%s%s.xls' % (name,time.strftime("%Y%m%d"))
    title=['岗位名称','地址','学历','单位','工资','地址','平均工资']
    index={}        #sheet的类目与数量
    jobExcList=[]   #有哪些经验类目sheet
    AverageList={}      #平均工资
    a=1
    file = xlwt.Workbook()
    while(a<11):
        set_response(name,a)
        a=a+1
    make_average()
    file.save(xlsName)
    make_echarts(ecahrtsName,index)
    print(jobExcList)
    print(index)








