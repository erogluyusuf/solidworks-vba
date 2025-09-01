' SOLIDWORKS VBA Macro
' Description: Otomatik olarak "YUSUF" yazılı bir anahtarlık oluşturur.
' -------------------------------------------------------------------

Option Explicit

Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swPart As SldWorks.PartDoc
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long

Sub main()

    ' --- Değiştirilebilir Parametreler ---
    Const ANAHTARLIK_UZUNLUK As Double = 0.06  ' Metre cinsinden (60 mm)
    Const ANAHTARLIK_GENISLIK As Double = 0.025 ' Metre cinsinden (25 mm)
    Const ANAHTARLIK_KALINLIK As Double = 0.003 ' Metre cinsinden (3 mm)
    Const DELIK_CAP As Double = 0.004           ' Metre cinsinden (4 mm)
    Const YAZI_METNI As String = "YUSUF"
    Const YAZI_DERINLIK As Double = 0.001       ' Metre cinsinden (1 mm)
    ' -----------------------------------------

    ' SOLIDWORKS uygulamasını al
    Set swApp = Application.SldWorks
    If swApp Is Nothing Then
        MsgBox "SOLIDWORKS çalışmıyor. Lütfen önce SOLIDWORKS'ü başlatın."
        Exit Sub
    End If

    ' Yeni bir parça dosyası oluştur
    Set swModel = swApp.NewPart
    If swModel Is Nothing Then
        MsgBox "Yeni parça dosyası oluşturulamadı."
        Exit Sub
    End If
    
    Set swPart = swModel

    ' -----------------------------------------------------
    ' 1. Adım: Anahtarlık Ana Gövdesini Oluşturma
    ' -----------------------------------------------------
    
    ' Ön Düzlem'i seç
    boolstatus = swModel.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
    
    ' Çizim başlat
    swModel.SketchManager.InsertSketch True
    
    ' Yuvalı dikdörtgen (slot) çiz. Bu, anahtarlık için ideal bir şekil.
    ' Toplam uzunluk = (ANAHTARLIK_UZUNLUK - ANAHTARLIK_GENISLIK) + ANAHTARLIK_GENISLIK
    Dim slotMerkezUzakligi As Double
    slotMerkezUzakligi = ANAHTARLIK_UZUNLUK - ANAHTARLIK_GENISLIK
    swModel.SketchManager.CreateCenteredLStraightSlot 0, 0, 0, slotMerkezUzakligi, ANAHTARLIK_GENISLIK
    
    ' Çizimi ekstrüzyon ile katıya dönüştür
    swModel.FeatureManager.FeatureExtrusion3(True, False, False, 0, 0, ANAHTARLIK_KALINLIK, 0.01, False, False, False, False, 1.74532925199433E-02, 1.74532925199433E-02, False, False, False, False, True, True, True, 0, 0, False)
    
    swModel.ClearSelection2 True
    
    ' -----------------------------------------------------
    ' 2. Adım: Anahtarlığın Başına Delik Açma
    ' -----------------------------------------------------

    ' Modelin ön yüzünü seç (koordinat ile seçim)
    boolstatus = swModel.Extension.SelectByID2("", "FACE", 0, 0, ANAHTARLIK_KALINLIK, False, 0, Nothing, 0)
    swModel.SketchManager.InsertSketch True
    
    ' Deliğin merkezini hesapla (sol kenardan biraz içeride)
    Dim delikMerkezX As Double
    delikMerkezX = -(ANAHTARLIK_UZUNLUK / 2) + (ANAHTARLIK_GENISLIK / 2)
    
    ' Delik için çember çiz
    swModel.SketchManager.CreateCircleByRadius delikMerkezX, 0, 0, DELIK_CAP / 2
    
    ' Ekstrüzyon ile kes (Tümünden)
    swModel.FeatureManager.FeatureCut4(False, False, False, 1, 0, 0.01, 0.01, False, False, False, False, 1.74532925199433E-02, 1.74532925199433E-02, False, False, False, False, False, True, True, True, True, False, 0, 0, False, False)

    swModel.ClearSelection2 True
    
    ' -----------------------------------------------------
    ' 3. Adım: Ortasına "YUSUF" Yazısını Yazma
    ' -----------------------------------------------------
    
    ' Tekrar modelin ön yüzünü seç
    boolstatus = swModel.Extension.SelectByID2("", "FACE", 0, 0, ANAHTARLIK_KALINLIK, False, 0, Nothing, 0)
    swModel.SketchManager.InsertSketch True
    
    ' Yazı için font ve boyut ayarları (İsteğe bağlı, varsayılanı kullanır)
    ' Bu kısmı değiştirmek için daha karmaşık kod gerekir, bu yüzden varsayılan font kullanılır.
    
    ' Yazıyı çizime ekle
    Dim swSketchText As SldWorks.SketchText
    Set swSketchText = swModel.SketchManager.InsertSketchText(YAZI_METNI, 0, 0, 0, 0, 100)
    
    ' Yazıyı ortalamak için (Eğer font ayarları yapılmadıysa yaklaşık bir konumlandırma)
    ' Gerekirse bu koordinatları değiştirin.
    Dim skText As SldWorks.SketchText
    Set skText = swModel.SketchManager.InsertSketchText(YAZI_METNI, -0.015, 0.004, 0, 0, 0)

    If Not skText Is Nothing Then
        ' Yazıyı biraz küçültmek için font yüksekliğini ayarla
        ' Bu kısım bazen versiyona göre çalışmayabilir. Manuel ayarlama gerekebilir.
        ' boolstatus = skText.SetHeight(0.008) ' 8mm yükseklik
    End If
    
    ' Yazıyı ekstrüzyon ile kes
    swModel.FeatureManager.FeatureCut4(True, False, False, 0, 0, YAZI_DERINLIK, 0.01, False, False, False, False, 1.74532925199433E-02, 1.74532925199433E-02, False, False, False, False, False, True, True, False, True, False, 0, 0, False, False)
    
    swModel.ClearSelection2 True
    
    ' Modeli ekrana sığdır
    swModel.ViewZoomtofit2
    
    MsgBox "Anahtarlık başarıyla oluşturuldu!"

End Sub
