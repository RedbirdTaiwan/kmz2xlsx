# -*- coding: UTF-8 -*-
import os, shutil, zipfile
import pandas as pd

def unzipkmz(f):
    with zipfile.ZipFile(f, 'r') as zip_ref:
        zip_ref.extractall('tmp')
    if os.path.isfile('tmp/doc.kml'):
        return True
    return False

def main():
    files = os.listdir('.')
    if not os.path.exists('tmp'):
        os.makedirs('tmp',755)

    for f in files:
        result = []
        if f[-3:].lower() == 'kmz':
            if unzipkmz(f):
                kml = open("tmp/doc.kml", "r", encoding='UTF-8')
                lines = kml.readlines()
                r = ['','']
                LatLonBox = []
                Polygon = ''
                Point = ''
                LineString = ''
                description = ''
                i = -1
                for line in lines:
                    # level
                    if line.find('<name>') > 0 :
                        j = line.find('<name>')-1
                        if j > i:
                            r.append(line[line.find('<name>')+6:line.find('</name>')])
                            i = j
                        else:
                            new_r = []
                            for i in range(0,len(r)):
                                new_r.append(r[i])
                            result.append(new_r)
                            
                            r[0] = ''
                            r[1] = ''
                            r = r[:j+2]
                            r.append(line[line.find('<name>')+6:line.find('</name>')])
                            i = j
                    
                    # LatLonBox
                    if line.find('<LatLonBox>') > 0 :
                        r[0] = "LatLonBox"
                        continue
                    if r and r[0] == "LatLonBox":
                        if line.find('<north>') > 0 :
                            LatLonBox.append(line[line.find('<north>')+7:line.find('</north>')])
                            continue
                        if line.find('<south>') > 0 :
                            LatLonBox.append(line[line.find('<south>')+7:line.find('</south>')])
                            continue
                        if line.find('<east>') > 0 :
                            LatLonBox.append(line[line.find('<east>')+6:line.find('</east>')])
                            continue
                        if line.find('<west>') > 0 :
                            LatLonBox.append(line[line.find('<west>')+6:line.find('</west>')])
                            continue
                    if r and len(LatLonBox) == 4 and r[0] == "LatLonBox":
                        r[1] = "N:%s,S:%s,E:%s,W:%s" %(LatLonBox[0],LatLonBox[1],LatLonBox[2],LatLonBox[3])
                        LatLonBox = []
                                        
                    # Polygon
                    if line.find('<Polygon>') > 0 :
                        r[0] = "Polygon"
                        continue
                    if r and r[0] == "Polygon":
                        if line.find('<coordinates>') > 0 :
                            Polygon = 'Polygon'
                            continue
                        if Polygon == 'Polygon' :
                            r[1] = line.replace('\t','').replace('\n','')
                            Polygon = ''
                            continue
                    
                    # Point
                    if line.find('<Point>') > 0 :
                        r[0] = "Point"
                        Point = 'Point'
                        continue
                    if Point == 'Point'and line.find('<coordinates>') > 0 :
                        r[1] = line.replace('\t','').replace('\n','').replace('<coordinates>','').replace('</coordinates>','')
                        Point = ''
                        continue

                    # LineString
                    if line.find('<LineString>') > 0 :
                        r[0] = "LineString"
                        continue
                    if r and r[0] == "LineString":
                        if line.find('<coordinates>') > 0 :
                            LineString = 'LineString'
                            continue
                        if LineString == 'LineString' :
                            r[1] = line.replace('\t','').replace('\n','')
                            LineString = ''
                            continue

                    # description
                    if line.find('<description>') > -1 :
                        description = line.replace('\t','').replace('\n','').replace('<description>','').replace('<![CDATA[','')
                        continue
                    if len(description) > 0 and line.find('</description>') < 0:
                        description += line.replace('\t','').replace('\n','')
                        continue
                    if line.find('</description>') > -1 :
                        description += line.replace(']]></description>','').replace('\t','').replace('\n','')
                        r.append(description)
                        description = ''
                        continue
                
                new_r = []
                for i in range(0,len(r)):
                    new_r.append(r[i])
                result.append(new_r)
                kml.close()
                pd.DataFrame(result).to_excel("%s.xlsx" %str(r[2]),index=False)
                shutil.rmtree('tmp')
                print("%s SUCCESS." %str(r[2]))

main()