Mən Kolin İnşaat şirkətində Lead Document Controller vəzifəsində çalışıram
Mənə sənəd göndərənlər:

1.  Bizim şirkətin departamentləri
    bizim şirkətin departamentlərinin göndərdiyi sənədlərin gəldiyi mail qovluğu: spp2dcc@kolin.com.tr Inbox/Sunulacaklar
    Bizim şirkətin departamentlərinin göndərdiyi sənədləraşağdakılar ola bilər:
    a. CLC, DWG, FRM, ITP, JSA, LOG, LST, MAR, MES, NCR, ORG, REP, SPE, SAR bu sənəd növləri Transmittal ilə göndərilir.
    Məsələn:

    - KLN-SPP2-FRM-MC-GN00-137_R00.xlsx - KLN-SPP2-FRM-MC-GN00-104_R03.pdf - KLN-SPP2-MAR-MC-GN00-145_R00.pdf - KLN-SPP2-MES-MC-GN00-003_R00.pdf
      b. SHD bu sənəd növləri Transmittal ilə göndərilir.
      bu sənədlər mənə bizi local file serverdə saxlanılır və link şəklində göndərilir.

          c. STQ bu sənəd növü adi mail ilə göndərilir.
          məsələn:
              - KLN-SPP2-STQ-MC-GN00-001_R00.xlsx
          d. LET bu sənəd növü adi mail ilə göndərilir.
          məsələn:
              - SPP2-KLN-PRO-LET-0001.docx


              Kodun etməli olduğu işlər:
              1. bunlar mənə mail əlavələri kimi Inbox/Sunulacaklar qovluğuna gəlir.
              2. 2 ci addımda yuxarıdakı qovluq scan edilir(həmən qovluqda yerləşən maillərin subjecti, bodysi, və əlvələrin adları).
              3. əgər a bəndində göstərilən sənəd növləri varsa aşağıdakı addımlar həyata keçirilir
                  step 1: \\10.10.8.253\DataServer\STP-S2-Projeler\Log\1. Outgoing\1. TRN qovluğunda növbəti transmittal folderi yaradılır.
                  step 2: növbəti transmittal folderinin içində 1. main, 2. attachments, 3. docs qovluqları yaradılır.
                  step 3: səndlər müvafiq olaraq 3. docs alt qovluğunun içində save edilir
                  step 4: bütün sənədlər KLN-SPP2-FRM-MC-GN00-137_R00 patterninə uyğun olmalıdır, yəni bununla birə bir eyni olmamalı, başlığı bu patternə uyğun
                          olmalıdır.
              4. Əgər b bəndində göstərilən sənəd növləri varsa aşağıdakı addımlar həyata keçirilir
                  step 1: step 1: \\10.10.8.253\DataServer\STP-S2-Projeler\Log\1. Outgoing\1. TRN qovluğunda növbəti transmittal folderi yaradılır.
                  step 2: növbəti transmittal folderinin içində 1. main, 2. attachments, 3. docs qovluqları yaradılır.
                  step 3: sənədlərin linkləri komputer tərəfindən açılmalı,  sənədlərin yerləşdiyi bir üst parent folder copyalanmalı və 3. docs alt qovluğunun içində save edilməldiri
                          məsələn:
                          \\DATA\DataServer\Elektrik\11- SHOPDRAWING\PROYAPI SUNUM\SOCKET SYSTEM INSTALLATION\ES03 bu link mailin bodysindədir
                          \\DATA\DataServer\Elektrik\11- SHOPDRAWING\PROYAPI SUNUM\SOCKET SYSTEM INSTALLATION a gəlib ES03 qovlugu kopyalanır və 3. docs alt qovluğunun içində save edilir.
                          \\DATA\DataServer\Elektrik\11- SHOPDRAWING\PROYAPI SUNUM\SOCKET SYSTEM INSTALLATION\EW13
                          \\DATA\DataServer\Elektrik\11- SHOPDRAWING\PROYAPI SUNUM\CABLE TRAY SYSTEM INSTALLATION\G13G
                          \\DATA\DataServer\Elektrik\11- SHOPDRAWING\PROYAPI SUNUM\CABLE TRAY SYSTEM INSTALLATION\GF05
                          bunlarda eyni qaydada mənə sənədlərin yox sənədlərin yerləşdiyi parent folderin kopyalanıb və 3. docs alt qovluğuna  save edilməsi lazımdr.
              5. Əgər c bəndində göstərilən sənəd növləri varsa aşağıdakı addımlar həyata keçirilir.
                  step 1: \\10.10.8.253\DataServer\STP-S2-Projeler\Log\1. Outgoing\3. STQ qovluğunda növbəti stq folderi yaradılır.
                  step 2: Nöbəti STQ folderi yaradılarkən diqqət edilməsi gərəkən məqamlar:
                      a. növbəti STQ folderinin rəqəmi tapılır
                          məsələn:
                          341. KLN-SPP2-STQ-CV-GN00-341 sonuncu SQT folderi budursa, növbəti yaradılacaq STQ folderinin rəqəmi 342 tapılır.

                      b. daha sonra mail əlavəsində olan STQ sənədinin kodu çıxarılır
                          məsələn:
                          KLN-SPP2-STQ-MC-GN00-001_R00.xlsx
                      c. daha sonra bu fayl başlığında olan kod KLN-SPP2-STQ-MC-GN00 bu hissəyə qədər kəsilir
                      d. daha sonra növbəti STQ folderinin rəqəmi. + fayl başlığından kəsilən hissə-növbəti STQ folderinin rəqəmi birləşdirilir və yeni STQ sənədinin adı yaradılır
                          məsələn:
                          342. KLN-SPP2-STQ-MC-GN00-342
                      e. daha sonra mail əlavəsindən götürülmüş xlsx formatında olan STQ sənədi və əlavə pdf ləri varsa, bu yeni yaradılmış STQ folderinin ichine save edilir.
                              342. KLN-SPP2-STQ-MC-GN00-342    bu folderin ichine

              6. Əgər d bəndində göstərilən sənəd növləri varsa aşağıdakı addımlar həyata keçirilir.
                      url: G:\My Drive\4-S1 ve S2 Ortak Dökümanlar\03-SPP LETTERS\SPP2-LET\1. KLN-PRO\01-Outgoing
                      step 1: yuxarıdakı url-ə daxil olunur
                      step 2: növbəti LET folderi yaradılır
                      step 3: mail əlavəsindən götürülmüş docx formatında olan LET sənədi bu yeni yaradılmış LET folderinin ichine save edilir.
                          məsələn:
                          son letter folderinin nömrəsi SPP2-KLN-PRO-LET-0086 dırsa, növbəti yaradılacaq letter folderi SPP2-KLN-PRO-LET-0087 olacaq.
                      step 4: yeni yaradılmış LET folderinin ichde 1. letter, 2. docs alt qovluqları yaradılır
                      step 5: mail əlavəsindən götürülmüş docx formatında olan LET sənədi 2. docs alt qovluğunun ichine save edilir.

2.  Proyapi

    göndərdiyi sənədlərin gəldiyi mail qovluğu: spp2dcc@kolin.com.tr Inbox/From Proyapi
    Proyapi göndərdiyi sənədləraşağdakılar ola bilər:
    a. TRN
    məsələn: - SPP2-PRO-KLN-TRN-0488 - SPP2-PRO-KLN-TRN-0487 - SPP2-PRO-KLN-TRN-0486 müxtəlif rəqəmlərlə ola bilər
    Bu maillər mənə DCC SPP2 | PROYAPI <dccspp2@proyapimusavirlik.com> ünvanıdan gəlir.
    b. STQ
    məsələn:
    KLN-SPP2-STQ-WE-GN00-309_R00_Prokon_Reply
    KLN-SPP2-STQ-MC-EW09-332_R00_Prokon_Reply və.s
    Bu maillər mənə DCC SPP2 | PROYAPI <dccspp2@proyapimusavirlik.com> ünvanıdan gəlir.

    c. LET
    məsələn:
    SPP2-PRO-KLN-LET-0020
    SPP2-PRO-KLN-LET-0019
    SPP2-PRO-KLN-LET-0018
    Bu maillər mənə DCC SPP2 | PROYAPI <dccspp2@proyapimusavirlik.com> ünvanıdan gəlir.

        Kodun etməli olduğu işlər:
        1. bunlar mənə mail əlavələri kimi Inbox/From Proyapi qovluğuna gəlir.
        2. Bu maillər mənə DCC SPP2 | PROYAPI <dccspp2@proyapimusavirlik.com> ünvanıdan gəlir.
        3. əgər a bəndində göstərilən sənəd növüdürsə bu zaman aşağıdakı addımdalar həyata keçirilir
            step 1: mail başlığı \\10.10.8.253\DataServer\QA-QC\QA-QC Proyapi\SPP2\99_Temporary\DCC\PRO-KLN-TRN qovluğunda axtarılır
            step 2: əgər mail başlığı orda tapılmırsa bu zaman mənim komputerimdə olan WP applicationdan +994 993 44 24 14 nömrəyə mesaj getməlidir.
                    Mesaj: "Səidə xanım, Yeni göndərdiyiniz transmittalları Data Serverə əlavə edə bilərsiniz, zəhmət olmasa?"
                    əgər mail başlığında olan koda uyğun gələn folder başlığı varsa bu zaman həmən folder copyalanmalı və \\10.10.8.253\DataServer\STP-S2-Projeler\Log\2. Incoming\1. TRN
                    qovluğuna yapışdırılmalıdır.
            step 3: yeni folder qeyd olunan incomin \\10.10.8.253\DataServer\STP-S2-Projeler\Log\2. Incoming\1. TRN qovluğuna kopyalandıqdan sorna içində 1. main, 2. attachments, 3. docs alt qovluqları yaradılmalıdır.
            step 4: qovluğun içində yerləşən zip folderi extract edilməlidir.
            step 5: extract edilmiş qovluq 3. docs alt qovluğunun içində save edilməlidir.
            step 6: 3. docs alt qovluğunun içində yerləşən save edilen qovluğun içindəki və ya alt qovluqlarındakı bütün pdf sənədləri 3. docs alt qovluğunun içinə kopyalanmalıdır.
            step 7: 3. docs alt qovluğuna kopyalanan bütün pdf sənədlərinin başlıqları KLN-SPP2-FRM-MC-GN00-137_R00 patterninə uyğun olmalıdır, yəni bununla birə bir eyni olmamalı, başlığı bu patternə uyğun
                    olmalıdır.
            step 8: daha sonra zip folder 2. attachments alt qovluğunun içinə köçürülməlidir.
            step 9: transmittal sənədinin pdf isə 1. main alt qovluğunun içində save edilməlidir.
            step 10: daha sonra bu mail pdf word formatına çevrilməlidir və 1. main alt qovluğunun içində save edilməlidir.
                        bunun uchun yazdigim powershell scripti de var
                                ##########################################   Incoming Trasnmittal   ################################################

        # Parent folder
        $parent = '\\10.10.8.253\DataServer\STP-S2-Projeler\Log\2. Incoming\1. TRN\SPP2-PRO-KLN-TRN-0489'

        # Yaradılacaq qovluqlar
        $folders = @(
            '1. main',
            '2. attachments',
            '3. docs'
        )

        # Parent yoxdursa, yarat
        if (-not (Test-Path -LiteralPath $parent)) {
            New-Item -ItemType Directory -Path $parent -Force | Out-Null
        }

        # Alt qovluqları yarat
        foreach ($f in $folders) {
            $p = Join-Path -Path $parent -ChildPath $f
            if (-not (Test-Path -LiteralPath $p)) {
                New-Item -ItemType Directory -Path $p -Force | Out-Null
            }
        }


        ##########################################   folder trasher   ################################################

        # Əsas qovluq ünvanını təyin edin
        $BaseFolder = "\\10.10.8.253\DataServer\STP-S2-Projeler\Log\2. Incoming\1. TRN\SPP2-PRO-KLN-TRN-0489"

        # Fayl və qovluq adlarını təyin edin (təkrar yazmamaq üçün)
        $FileName = "SPP2-PRO-KLN-TRN-0489"
        $PdfFile = "$BaseFolder\$FileName.pdf"
        $ZipFile = "$BaseFolder\$FileName.zip"
        $ExtractedFolder = "$BaseFolder\$FileName" # Zip çıxarıldıqdan sonra yaranacaq qovluq
        $MainFolder = "$BaseFolder\1. main"
        $AttachmentsFolder = "$BaseFolder\2. attachments"
        $DocsFolder = "$BaseFolder\3. docs"

        Write-Host "Əməliyyatlara başlanır..."

        # 1. SPP2-PRO-KLN-TRN-0489.pdf faylını 1. main qovluğuna KÖÇÜRMƏK (Move-Item)
        # Qeyd: Bu əməliyyat faylı orijinal yerindən silir.
        if (Test-Path $PdfFile) {
            Move-Item -Path $PdfFile -Destination $MainFolder -Force
            Write-Host "1. $FileName.pdf faylı '$($MainFolder)' qovluğuna KÖÇÜRÜLDÜ." -ForegroundColor Green
        } else {
            Write-Host "1. XƏTA: $PdfFile tapılmadı." -ForegroundColor Yellow
        }

        # 2. SPP2-PRO-KLN-TRN-0489 nömrəli zipi extract etmək
        if (Test-Path $ZipFile) {
            try {
                Expand-Archive -Path $ZipFile -DestinationPath $BaseFolder -Force
                Write-Host "2. $FileName.zip faylı '$($BaseFolder)' qovluğuna çıxarıldı."

                # 3. Zipdən çıxarılmış SPP2-PRO-KLN-TRN-0489 qovluğunu 3. docs folderinə köçürmək
                if (Test-Path $ExtractedFolder) {
                    Move-Item -Path $ExtractedFolder -Destination $DocsFolder -Force
                    Write-Host "3. Çıxarılmış '$($FileName)' qovluğu '$($DocsFolder)' qovluğuna KÖÇÜRÜLDÜ."
                } else {
                    Write-Host "3. XƏTA: Çıxarılmış qovluq ($ExtractedFolder) tapılmadı. Ola bilsin ki, zip çıxarıldıqda fərqli adla qovluq yaranıb." -ForegroundColor Yellow
                }

                # 4. SPP2-PRO-KLN-TRN-0489 nömrəli zipi 2. attachments qovluğuna köçürmək
                Move-Item -Path $ZipFile -Destination $AttachmentsFolder -Force
                Write-Host "4. $FileName.zip faylı '$($AttachmentsFolder)' qovluğuna KÖÇÜRÜLDÜ."

            } catch {
                Write-Host "2/3/4. XƏTA: Zip əməliyyatlarında problem yarandı. $_" -ForegroundColor Red
            }
        } else {
            Write-Host "2/3/4. XƏTA: $ZipFile tapılmadı." -ForegroundColor Yellow
        }

        Write-Host "Bütün əməliyyatlar tamamlandı."


        ###################################################### incoming doucments to parent folder ######################################################

        $source = "\\10.10.8.253\DataServer\STP-S2-Projeler\Log\2. Incoming\1. TRN\SPP2-PRO-KLN-TRN-0489\3. docs"
        $destination = $source  # parent folder is itself since PDFs go to 3. docs
        Get-ChildItem -Path $source -Recurse -Filter *.pdf | ForEach-Object {
            $target = Join-Path $destination $_.Name
            if (Test-Path $target) {
                $basename = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
                $ext = $_.Extension
                $newName = "$basename" + "_copy" + $ext
                $target = Join-Path $destination $newName
            }
            Copy-Item -Path $_.FullName -Destination $target
        }

        ###############################       -R i _R ə çevirən kod        ###########################################

        $folder = "\\10.10.8.253\DataServer\STP-S2-Projeler\Log\2. Incoming\1. TRN\SPP2-PRO-KLN-TRN-0489\3. docs"

        Get-ChildItem -Path $folder -Filter *.pdf | ForEach-Object {
            $oldName = $_.Name
            $newName = $oldName -replace '-R', '_R'
            if ($oldName -ne $newName) {
                Rename-Item -Path $_.FullName -NewName $newName
            }
        }



        ################################      Title Maker      ###########################################################################################################################

        # Əməliyyat ediləcək qovluq ünvanını təyin edin
        $TargetFolder = "\\10.10.8.253\DataServer\STP-S2-Projeler\Log\2. Incoming\1. TRN\SPP2-PRO-KLN-TRN-0489\3. docs"

        Write-Host "Qovluqdakı faylların adlarının dəyişdirilməsinə başlanır: $TargetFolder"

        # Qovluqdakı bütün PDF fayllarını tapın
        $Files = Get-ChildItem -Path $TargetFolder -Filter "*.pdf" -File

        # Hər bir PDF faylı üzərində dövr edin
        foreach ($File in $Files) {
            # Mövcud fayl adını (uzantısız) əldə edin
            $BaseName = $File.BaseName

            # REGEX istifadə edərək təmizləmə əməliyyatını edin:
            # Pattern: (_R[0-9]{2})_.*
            # 1. (_R[0-9]{2}): "_R" ilə başlayan və ardınca iki rəqəm gələn hissəni tapır və saxlayır (Group 1).
            # 2. _.* : Bu hissədən sonra gələn alt xətt "_" və ardınca gələn hər şeyi tapır (bu hissə silinəcək).

            # Fayl adını formatınıza uyğunlaşdırmaq üçün dəyişdirin.
            $NewBaseName = $BaseName -replace '(_R[0-9]{2})_.*', '$1'

            # Yeni fayl adını (uzantısı ilə birlikdə) yaradın
            $NewFileName = "$NewBaseName.pdf"
            $NewPath = Join-Path -Path $TargetFolder -ChildPath $NewFileName

            # Əgər dəyişiklik varsa və yeni fayl adı hələ yoxdursa, adı dəyişdirin
            if ($BaseName -ne $NewBaseName) {
                if (-not (Test-Path $NewPath)) {
                    Rename-Item -Path $File.FullName -NewName $NewFileName -Force
                    Write-Host "Dəyişdirildi: '$($File.Name)' -> '$($NewFileName)'" -ForegroundColor Green
                } else {
                    Write-Host "XƏTA: '$($NewFileName)' adlı fayl artıq mövcuddur. '$($File.Name)' atlandı." -ForegroundColor Yellow
                }
            }
        }

        Write-Host "Bütün əməliyyatlar tamamlandı."



        ############################################################    Convertor to Word       ##########################################################################


        # Əsas dəyişənləri təyin edin
        $PDFPath = "\\10.10.8.253\DataServer\STP-S2-Projeler\Log\2. Incoming\1. TRN\SPP2-PRO-KLN-TRN-0489\1. main\SPP2-PRO-KLN-TRN-0489.pdf"
        $DOCXPath = $PDFPath -replace '\.pdf$', '.docx' # Eyni adda, lakin .docx uzantısı ilə

        Write-Host "Çevrilmə prosesinə başlanır..."

        # Microsoft Word tətbiqini işə salın
        $Word = New-Object -ComObject Word.Application
        $Word.Visible = $false # Word tətbiqini göstərməyin

        try {
            # 1. PDF faylını açın (Word bunu avtomatik olaraq çevirir)
            $Document = $Word.Documents.Open($PDFPath)
            Write-Host "PDF sənədi Word tərəfindən açıldı və çevrildi..."

            # 2. Sənədi DOCX formatında (wdFormatDocumentDefault = 16) qeyd edin
            # Qeyd: Word 2007 və sonrakı versiyaları üçün 16 dəyəri .docx formatını ifadə edir.
            $Document.SaveAs([ref]$DOCXPath, [ref]16)
            Write-Host "Sənəd müvəffəqiyyətlə DOCX formatında qeyd edildi: $DOCXPath" -ForegroundColor Green

            # 3. Sənədi bağlayın
            $Document.Close()

        } catch {
            Write-Host "XƏTA: Faylın çevrilməsi zamanı problem yarandı. Microsoft Word quraşdırılmamış ola bilər və ya Word bu PDF-i aça bilmədi." -ForegroundColor Red
            Write-Host "Xəta mesajı: $($_.Exception.Message)"
        } finally {
            # Həmişə Word tətbiqini bağlayın və resursları təmizləyin
            if ($Word) {
                $Word.Quit()
            }
            # COM obyektini yaddaşdan təmizləyin (bu vacibdir)
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | Out-Null
            Remove-Variable Word -ErrorAction SilentlyContinue
        }

        Write-Host "Proses tamamlandı."

3.  əgər b bəndində göstərilən sənəd növüdürsə bu zaman aşağıdakı addımdalar həyata keçirilir
    step 1: \\10.10.8.253\DataServer\STP-S2-Projeler\Log\1. Outgoing\3. STQ qovluğunda axtarılır və uyğun gələn folder tapılır
    məsələn:
    KLN-SPP2-STQ-WE-GN00-309_R00_Prokon_Reply
    step 2: daha sonra file bashligi bu formata KLN-SPP2-STQ-WE-GN00-309 salinir
    step 3: daha sonra yuxarıda təmin edilmish qovlugun ichinde bu koda uygun gelen folder tapilir KLN-SPP2-STQ-WE-GN00-309
    step 4: daha sonra file bashligi bu formatdan (KLN-SPP2-STQ-WE-GN00-309_R00_Prokon_Reply) bu formata (KLN-SPP2-STQ-WE-GN00-309_R00 Reply) salinir
    step 5: daha sonra tapilan folderin ichine save edilir

4.  əgər c bəndində göstərilən sənəd növüdürsə bu zaman aşağıdakı addımdalar həyata keçirilir
    step 1: url: G:\My Drive\4-S1 ve S2 Ortak Dökümanlar\03-SPP LETTERS\SPP2-LET\1. KLN-PRO\02-Incoming qovlugunda növbəti LET folderi yaradılır.
    məsələn:
    son letter folderinin nömrəsi SPP2-PRO-KLN-LET-0021 dırsa, növbəti yaradılacaq letter folderi SPP2-PRO-KLN-LET-0022 olacaq.
    step 2: daha sonra qovluğun içində 1. letter, 2. docs alt qovluqları yaradılır
    step 3: mail əlavəsindən götürülmüş pdf formatında olan LET sənədi 1. letter alt qovluğunun ichine save edilir.
    step 4: əhər əlavə fayllar varsa onlar 2. docs alt qovluğunun ichine save edilir.
