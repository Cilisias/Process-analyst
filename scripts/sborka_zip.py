import os.path
import zipfile
from timeit import default_timer as timer
from shutil import copyfile

startTimer = timer()

textfile = r'C:\Users\KarklinaEA\Desktop\Имущественный комплекс КП\Охранные зоны ГГС_ГНС\ГНС_10.07.2023\для сборки ГНС_10.07.2023.txt'
pathToZipFolder = r'C:\Users\KarklinaEA\Desktop\Имущественный комплекс КП\Охранные зоны ГГС_ГНС\ГНС_10.07.2023\ZIP'
dirname = r'C:\Users\KarklinaEA\Desktop\Имущественный комплекс КП\Охранные зоны ГГС_ГНС\ГНС_10.07.2023'
dirnameXML = r'C:\Users\KarklinaEA\Desktop\Имущественный комплекс КП\Охранные зоны ГГС_ГНС\ГНС_10.07.2023\XML'
dst = r'C:\Users\KarklinaEA\Desktop\Имущественный комплекс КП\Охранные зоны ГГС_ГНС\ГНС_10.07.2023\ZIP'

numb_list = []
with open(textfile, 'r', encoding='utf-8') as f:
    for item in f:
        numb_list.append([str(i) for i in item.replace('\n', '').strip().split(';')])

    #print(numb_list)

for i in range(len(numb_list)):
    newzip = zipfile.ZipFile(os.path.join(dst,  numb_list[i][1].replace('.xml', '.zip')), 'w')  #create ZIP file
    if numb_list[i][3].find('|'):
        applf = numb_list[i][3].split('|')
        for item in range(len(applf)):
            newzip.write(os.path.join(dirname, applf[item]), arcname=os.path.join('', applf[item]))  #file
            newzip.write(os.path.join(dirname, applf[item] + '.sig'), arcname=os.path.join('', applf[item] + '.sig'))  #file

    else:
        newzip.write(os.path.join(pathToZipFolder, numb_list[i][3]), arcname=os.path.join('', numb_list[i][3]))  # file
    newzip.write(os.path.join(dirname,'AppliedFiles//' + numb_list[i][0] + '.pdf'), arcname=os.path.join('AppliedFiles', numb_list[i][0] + '.pdf'))  #file
    newzip.write(os.path.join(dirname,'AppliedFiles//' + numb_list[i][0] + '.pdf' + '.sig'), arcname=os.path.join('AppliedFiles', numb_list[i][0] + '.pdf' + '.sig'))  #file
    newzip.write(os.path.join(dirnameXML, numb_list[i][1]), arcname=os.path.join('', numb_list[i][1]))  #file in folder
    newzip.write(os.path.join(dirnameXML, numb_list[i][2]), arcname=os.path.join('', numb_list[i][2]))  #file in folder
    newzip.write(os.path.join(dirnameXML, numb_list[i][1] + '.sig'), arcname=os.path.join('', numb_list[i][1] + '.sig'))  #file in folder
    newzip.write(os.path.join(dirnameXML, numb_list[i][2] + '.sig'), arcname=os.path.join('', numb_list[i][2] + '.sig'))  #file in folder
    newzip.close()  #close ZIP file

    print(newzip.filename)

endTimer = timer()

print("Time taken: ", endTimer - startTimer)