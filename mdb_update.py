import pandas as pd
import pyodbc
import xlsxwriter
import sqlalchemy
import os

def clear():
  os.system('cls||clear')

def connect_to_database():
    try:
        con = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\moustapha\Desktop\\MDB_database.mdb;')
        return True
    except:
        clear()
        return False
    
con = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\moustapha\Desktop\\MDB_database.mdb;')
cursor = con.cursor()

excel_file = "updated.xlsx"
row = 200

def readme():
    clear()
    print("""
    DİKKAT :
    --------
    Programı orjinal veritabanı dosyasında kullanmayınız ve lütfen yedek/kopya dosyalar üzerinde işlem yapınız.
    Kullanımdan dolayı oluşabilecek tüm olumsuz koşullardan ve hatalardan kullanıcı sorumludur.
    Yazılımı hazırlayan kişi, hiç bir suretle oluşabilecek hata ve zararlardan sorumlu tutulamaz.
    Programın kullanımıyla, bu alandaki tüm koşullar kabul edilmiş sayılır.
    
    Program 'product', 'product_code' ve 'product_price' olarak 3 alan içeren bir veritabanı için hazırlanmıştır.
    
    Veritabanında Primary Key alanı varsa, programın doğru şekilde çalışması için değişiklik yapılması gereklidir.
    
    Program adımlarının / menülerinin -aynı oturum için- sadece 1 kez çalıştırılması gerekmektedir. 
    
    Aynı oturumda program adımlarının/menülerinin tekrar çalıştırılması hataya yol açacaktır.
    
    KULLANIM :
    ----------
    - Programın MDB veritabanını Excel dosyasına aktarabilmesi için veritabanının program ile aynı dizinde olması gereklidir.
    - Varsayılan veritabanı adı: MDB_database.mdb ve okunan Table adı: Table_Name
    varsayılan Excel dosyası adı: updated.xlsx ve Sheet adı: Sheet1'dir.
    - 'row' Değişkeni ile, Excel dosyasından Veritabanına yazma esnasında 1'den 200. satıra ve 3 kolona kadar bilgi okunarak yazılır.
    Yani Table'ın 200 satır, 3 kolon veri içerdiği varsayılmıştır.
    Programın işleme alacağı MDB dosyası veritabanı alanları örnek olarak;
    
       | product | product_code | product_price
       ----------+--------------+--------------
       | sample1 |   00000001   | 100,50      |
       ----------+--------------+--------------
       | sample2 |   00000002   | 200,75      |
       ----------+--------------+--------------      şeklindedir.
    
    - Program, güncelleme yapılan Excel dosyasındaki veriyi -güvenlik ve olası hataları önlemek maksadıyla- veritabanında 'new_table' adında
    yeni oluşturduğu bir Table üzerine yazar. Kullanıcı, MDB veritabanını kullanmadan önce Table ile ilgili isim, alan vb. değişiklikleri
    gerçekleştirmelidir.
    
    İyi çalışmalar, çalışmalarınızda başarılar.    
          """)
    input("\n\tDevam etmek için [Enter] ")

def mdb_to_excel():
    clear()
    if (connect_to_database()) == True:
        while True:
            try:
                choice = input("\n\tMDB veritabanını Excel dosyasına aktarma işlemi E / H ?")
                if choice.upper() in ["E" , "e"]:  
                    df = pd.read_sql("SELECT * FROM Table_Name", con)
                    with pd.ExcelWriter("updated.xlsx") as writer:
                        df.to_excel(writer, sheet_name="Sheet1")
                    clear()
                    input("\n\tİşlem başarılı. Lütfen Excel dosyasını kontrol ediniz. Devam etmek için [Enter] ")
                    break
                else:
                    break
            except:
                clear()
                input("\n\tDönüştürme işleminde hata alındı. LÜtfen MDB ve Excel dosyalarını kontrol edin. Devam etmek için [Enter] ")
                break
    else:
        print("\n\tVeritabanı bağlantısı başarısız. Lütfen programı tekrar çalıştırın. devam etmek için [Enter] ")

def excel_to_mdb(row):
    while True:
        clear()
        try:
            choice = input("\n\tExcel dosyasından MDB veritabanına aktarma işlemi E / H ?")
            if choice.upper() in ["E" , "e"]:
                cursor.execute("CREATE TABLE new_table(product TEXT, product_code INT, product_price FLOAT)")
                df = pd.read_excel(excel_file)
                for i in range(0,row):
                    cursor.execute("INSERT INTO new_table(product, product_code, product_price) VALUES(?,?,?)" ,(df.iat[i,1], int((df.iat[i,2])), float(df.iat[i,3]) ))
                    con.commit()
                clear()
                input("\n\tişlem tamamlandı. Lütfen Veritabanını kontrol ediniz. Bağlantı kapatılıyor. Devam etmek için [Enter] ")
                con.close()
                break
            else:
                break
        except:
            clear()
            input("\n\tOkuma / Yazma işleminde hata alındı. LÜtfen MDB ve Excel dosyalarını kontrol edin. Devam etmek için [Enter] ")
            break

if (connect_to_database()) == True:
    while True:
        clear()
        print("\n\tMDB VERİTABANI GÜNCELLEME\n")
        print("\t"+"*"*25)     
        print("\n\tKullanmadan önce Readme bölümünü okuyunuz !\n\n\tProgram adımlarını -gerekli kontrolleri yaptıktan sonra- sadece 1 kez çalıştırınız.\n")
        print("\t(1) MDB Veritabanını Excel Dosyasına dönüştür\n\n\t(2) Excel Dosyasını MDB Veritabanına Yaz\n\n\t(3) Readme\n\n\t(4) Çıkış")
        choice = input("\n\tSeçiminiz: ")
        
        if (choice == "1"):
            mdb_to_excel()
        elif (choice == "2"):
            excel_to_mdb(row)
        elif (choice == "3"):
            readme()
        elif (choice == "4"):
            try:
                con.close()
                print("\n\tVeritabanı bağlantısı sonlandırıldı. Çıkış yaptınız.")
                break
            except:
                print("\n\tÇıkış yaptınız.")
                break
        else:
            input("\n\tSeçenekler arasından seçim yapılmadı. Devam etmek için [Enter] ")
else:
    clear()
    print("\n\tVeritabanı bağlanı işlemi başarısız. Lütfen Veritabanı dizinini, ismini ve uzantısını kontrol edin.")
    input("\n\tHata alınarak programdan çıkış yapıldı. Readme dosyasındaki yönergeleri kontrol ederek tekrar deneyin.Devam etmek için [Enter] \n")
