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