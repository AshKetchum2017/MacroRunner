# MacroRunner

MacroRunner adalah macro VBA untuk CorelDRAW yang menyusun beberapa project GMS ke dalam antrean, menjalankannya satu per satu, dan mencatat status serta durasi setiap langkah.

Satu konfigurasi mempunyai `ConfigID`, urutan macro, dan `MacroBehavior` opsional. Tanpa Behavior, setiap form dibuka untuk dioperasikan secara manual. Dengan Behavior, runner dapat mengirim instruksi ke macro tujuan yang sudah menyediakan kontrak integrasi.

## Main Features

### Macro Queue

`MacroRunnerMenu` membuka `MacroQueue` melalui **Queue Settings**. Di dalam antrean:

- **Set** membuat konfigurasi baru melalui `MacroSelection`.
- **Modify** membuka konfigurasi yang dipilih dan menyimpan perubahan pada baris yang sama.
- **Remove** menghapus satu konfigurasi setelah konfirmasi.
- **Clear** menghapus semua konfigurasi setelah konfirmasi; tindakan ini tidak dapat dibatalkan dari UI.
- **Select** memilih konfigurasi untuk dijalankan.

Daftar ditampilkan sebagai `ConfigID | MacroA > MacroB`. `ConfigID` wajib diisi saat **Save**, tetapi tidak harus menjadi nama project GMS. Konfigurasi lama yang dimigrasikan tanpa `ConfigID` masih dapat dibaca; isi `ConfigID` saat menyunting dan menyimpannya kembali.

Di `MacroSelection`, pilih GMS dari `cmbMacroLists` lalu tekan **Add**, atau ketik nama file GMS tanpa `.gms` langsung di `txbSelectedMacro`. Pisahkan nama dengan titik koma:

```text
ExportRelated; AutoSaveNCreate;
```

Macro yang sama boleh muncul lebih dari sekali. Menghapus tepat satu karakter `;` di `txbSelectedMacro` juga menghapus token macro di depannya; perubahan teks lain mengikuti perilaku editor biasa. Jika pencarian folder GMS gagal, nama masih dapat diketik manual. Ketersediaan project baru diperiksa saat **Process**.

### Process, Continue, dan Statistics

1. Pilih konfigurasi di `MacroQueue`, lalu tekan **Select**.
2. Tekan **Process**. Runner memeriksa seluruh antrean sebelum menjalankan langkah pertama.
3. Setelah form utama suatu macro ditutup dan langkahnya menjadi `OK`, tekan **Continue** untuk menjalankan macro berikutnya.
4. Lihat hasil di `lbxStatistics` atau tekan **Copy Statistic** untuk menyalin seluruh statistik ke clipboard.

Status langkah adalah `PENDING`, `RUNNING`, `OK`, `ERROR`, dan `SKIPPED`. Jika sebuah langkah gagal, antrean berhenti dan langkah berikutnya menjadi `SKIPPED`. Durasi ditampilkan dalam detik dengan tiga angka desimal; waktu menunggu tombol **Continue** tidak dihitung. Penutupan form utama, termasuk waktu interaksi pengguna, menentukan akhir langkah. `MacroRunnerMenu` tidak dapat ditutup saat macro sedang berjalan.

## MacroBehavior

Isi `txbMacroBehavior` untuk menjalankan instruksi pada form target yang mendukungnya. Enter membuat baris baru, sedangkan Tab tanpa modifier menyisipkan lima spasi. Textbox yang kosong atau hanya berisi spasi mempertahankan alur manual; `Form[]` adalah Behavior eksplisit dan tetap memerlukan validator di target.

Format dasar:

```text
{MacroID:FormName[Target=Value;@Action];FormName[Target=Value]}
```

Satu blok `{...}` mewakili **satu kemunculan** macro di `txbSelectedMacro`, dalam urutan yang sama. Jika sebuah macro muncul dua kali, tulis dua blok terpisah. Antarblok dapat dipisahkan dengan spasi atau baris baru.

| Sintaks | Makna |
| --- | --- |
| `MacroID` | Nama file GMS tanpa `.gms` |
| `FormName[...]` | Tahap pada form target; nama mengikuti `(Name)` di VBE |
| `Target=Value` | Mengisi kontrol atau properti yang didaftarkan oleh target |
| `@Action` | Menjalankan action yang didaftarkan oleh target, tanpa argumen |
| `;` | Pemisah instruksi atau tahap form |

Spasi, tab, dan baris baru boleh digunakan di luar string, tetapi jangan di tengah identifier. Nama macro, form, target, action, dan keyword tidak peka kapitalisasi. Isi di dalam tanda kutip dipertahankan. Parser MacroRunner hanya memeriksa grammar umum; setiap macro tujuan menentukan sendiri form, target, action, dan urutan operasi yang sah.

### Nilai yang dapat ditulis

| Bentuk | Makna |
| --- | --- |
| `"Teks"` | String; token seperti `"{A}1,2,3"` diteruskan ke macro tujuan |
| `"Label ""A"""` | String `Label "A"` dengan tanda kutip ganda di dalamnya |
| `True`, `False` | Boolean |
| `123`, `1.5` | Number; desimal memakai titik |
| `Default` | Nilai bawaan yang pemetaan dan ketersediaannya ditentukan target |
| `Nothing`, `Empty`, `Null` | Lewati assignment, setelah targetnya tetap divalidasi |

`"Default"` adalah string biasa, bukan keyword. `""` mengisi string kosong, sedangkan `Empty` membiarkan nilai target. Backslash di dalam string Behavior adalah karakter biasa. DSL ini tidak mengevaluasi ekspresi VBA atau mengizinkan pemanggilan prosedur bebas.

### Contoh: ExportRelated

Contoh berikut memakai nama target dan action dari kontrak **ExportRelated** yang didokumentasikan untuk integrasi tersebut. Pastikan versi ExportRelated yang terpasang memang menyediakan `ValidateBehavior` dan `RunBehavior` yang cocok.

`txbConfigID`:

```text
Export PDF per halaman
```

`txbSelectedMacro`:

```text
ExportRelated;
```

`txbMacroBehavior`:

```text
{ExportRelated:
    ExportRelatedMenu[@cmdAddSetting];
    ExportRelatedSettings[
        txbName="{A}1,2,3";
        txbPage="{1-3}";
        cmbExFormat=".pdf";
        chkParentDirectory=False;
        txbDirectory=Default;
        @cmdSave
    ];
    ExportRelatedMenu[
        lbxSettingLists.Index=1;
        chkLayer1=True;
        chkLayer2=False;
        chkLayer3=False;
        @cmdExport;
        @cmdClose
    ]
}
```

Di sini `{1-3}` pada `txbPage` menunjuk keluaran per halaman sesuai aturan token ExportRelated, sementara `{A}1,2,3` pada `txbName` diteruskan apa adanya ke parser nama target. `txbDirectory=Default` memerlukan direktori bawaan yang valid pada instalasi ExportRelated. Periksa kecocokan jumlah nama, halaman, dokumen, layer, dan lokasi output sebelum menjalankan ekspor. `lbxSettingLists.Index=1` memilih item pertama; `.Index` adalah target khusus DSL, bukan nama properti MSForms yang bisa diganti menjadi `.ListIndex`.

Daftar token dan target ExportRelated dapat berubah di project tersebut. Untuk token yang belum tercatat di panduan, periksa implementasi dan kontrak target yang sedang dipasang; MacroRunner tidak menafsirkan token nama atau halaman di dalam string.

### Save dan preflight

**Save** mewajibkan `ConfigID`, memeriksa urutan macro dan struktur Behavior, lalu menyimpan teks Behavior aslinya. Pada tahap ini runner belum menguji apakah target seperti `txbName` atau `@cmdExport` didukung, apakah file GMS tersedia, atau apakah nilai `Default` dapat dipakai.

**Process** membaca ulang konfigurasi dan memeriksa semua project GMS. Untuk setiap blok Behavior eksplisit, runner memanggil `MRTargetBridge.ValidateBehavior` milik project target sebelum menjalankan langkah pertama. Jika validasi salah satu langkah gagal, antrean tidak mulai berjalan. Saat eksekusi, `MRTargetBridge.RunBehavior` target menjalankan instruksi; runner memerlukan callback konfirmasi dari bridge. Pemeriksaan awal tidak menjamin kondisi dokumen atau file output tetap sama ketika langkah benar-benar berjalan.

Behavior kosong tidak memerlukan validator Behavior, tetapi project tujuan tetap harus memiliki bridge pembuka form dan hook penutupan. Jangan memakai blok `Form[]` sebagai pengganti Behavior kosong untuk mode manual.

## Penyimpanan konfigurasi

Konfigurasi disimpan sebagai JSON UTF-8 versi 2 di:

```text
%APPDATA%\RinCorelMacros\MacroRunner\Configurations.json
```

Setiap entri menyimpan `configId`, `sequence`, dan `behavior`:

```json
{"version":2,"configurations":[{"configId":"Manual produksi","sequence":"ExportRelated;","behavior":""}]}
```

Jika file JSON belum ada, runner membaca konfigurasi lama dari registry `RinCorelMacros/MacroRunner/SequencesV1`. Perubahan pertama melalui **Save**, **Remove**, atau **Clear** menulisnya ke JSON; registry lama tidak dihapus. Setelah JSON ada, file tersebut menjadi sumber konfigurasi. Konfigurasi lama tanpa `configId` bisa dimuat, tetapi **Save** berikutnya tetap memerlukan `ConfigID`. Pilihan aktif dan statistik hanya berlaku selama sesi runner.

## Integrasi di CorelDRAW/VBE

Source di `src/classes/` adalah class module; pertahankan `(Name)` masing-masing saat memasangnya di project MacroRunner. File di `src/forms/` adalah code-behind untuk UserForm dengan nama yang sesuai, bukan file desain `.frm` lengkap. Siapkan kontrol yang disebut di kode form, termasuk `cmdContinue` pada `MacroRunnerMenu` serta `txbConfigID` dan `txbMacroBehavior` pada `MacroSelection`.

MacroRunner mencari `.gms` di folder GMS milik profil CorelDRAW aktif melalui `Application.UserDataPath`. Project tujuan harus sudah dimuat oleh CorelDRAW. `MRCatalog` mempunyai registrasi form untuk `ExportRelated` dan `AutoDistributeUF` serta beberapa nama prosedur pembuka lain; keberadaan nama dalam daftar itu tidak otomatis memasang bridge atau menjamin kompatibilitas target.

Untuk alur manual, pasang `src/Integration/MRTargetBridge.bas` pada project GMS tujuan dan gabungkan `FormLifecycleSnippet.vba` ke UserForm utamanya. Form harus menyediakan `MRBindRunner`, `MRDetachRunner`, dan callback saat `UserForm_Terminate`, agar runner mengetahui kapan langkah selesai. Sesuaikan integrasi dengan lifecycle form tujuan; bridge generik menolak form yang sudah terbuka. `AutoDistributeUF` memakai bridge khusus dari project AutoDistributeUF.

Untuk Behavior eksplisit, target juga harus menyediakan implementasi `MRTargetBridge.ValidateBehavior` dan `MRTargetBridge.RunBehavior`, kontrak form/target/action, serta callback konfirmasi yang diharapkan `MRPresenter`. Template bridge generik di repo ini hanya menyediakan `OpenMacro`; integrasi Behavior untuk ExportRelated berada di project targetnya. Jika parser atau protokol Behavior diperbarui, perbarui salinan source bersama di MacroRunner dan target yang menggunakannya. Setelah pemasangan, jalankan **Debug > Compile** pada project terkait dan uji alur manual serta Behavior dengan dokumen contoh.

## Struktur source

| Lokasi | Peran |
| --- | --- |
| `src/forms/` | `MacroRunnerMenu`, `MacroQueue`, `MacroSelection` |
| `src/classes/MRPresenter.cls` | Koordinasi UI, preflight, eksekusi, dan callback |
| `src/classes/MRBehaviorParser.cls` | Parser grammar Behavior |
| `src/classes/MRCatalog.cls` | Pencarian GMS dan resolusi project/form |
| `src/classes/MRRunModel.cls` | Status antrean dan statistik |
| `src/classes/MRSequenceStore.cls`, `MRConfigJson.cls` | Penyimpanan dan pembacaan konfigurasi |
| `src/Integration/` | Template bridge dan contoh hook lifecycle untuk target |

## Lisensi dan masukan

Project ini menggunakan lisensi MIT; lihat [LICENSE](LICENSE).

Source boleh dipelajari dan dikembangkan, dan issue/feedback tentang bug, edge case, CorelDRAW API, architecture, atau improvement sangat dihargai.
