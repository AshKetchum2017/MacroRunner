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
- **Up** / **Down** memindahkan konfigurasi terpilih satu posisi dalam antrean. Urutan baru langsung disimpan; urutan macro *di dalam* konfigurasi tidak berubah.
- **Select** memilih konfigurasi untuk dijalankan.
- **Close** menutup tampilan Queue tanpa mengakhiri sesi MacroRunner.

Setiap konfigurasi tampil sebagai satu blok dua baris dalam `frmMacroLists`: `ConfigID` di atas dan urutan macro di bawahnya, misalnya:

```text
Produksi label
ExportRelated > AutoSaveNCreate
```

Klik salah satu baris untuk memilih blok; pilihan tetap mengikuti konfigurasi saat dipindahkan dengan **Up** / **Down**. Tombol perpindahan hanya aktif jika ada posisi tujuan. `ConfigID` wajib diisi saat **Save**, tetapi tidak harus menjadi nama project GMS. Konfigurasi lama yang dimigrasikan tanpa `ConfigID` tetap dapat dibaca dan hanya menampilkan urutan macro; isi `ConfigID` saat menyunting dan menyimpannya kembali.

`MacroQueue` dapat dibuka kembali dari **Queue Settings** selama runner siap menerima pengaturan. **Select**, **Close**, dan tombol tutup jendela menyembunyikan Queue; ketiganya tidak mengakhiri MacroRunner. Tutup Queue dahulu sebelum memakai **Process** atau menutup `MacroRunnerMenu`.

Di `MacroSelection`, pilih GMS dari `cmbMacroLists` lalu tekan **Add**, atau ketik nama file GMS tanpa `.gms` langsung di `txbSelectedMacro`. Pisahkan nama dengan titik koma:

```text
ExportRelated; AutoSaveNCreate;
```

Macro yang sama boleh muncul lebih dari sekali. Menghapus tepat satu karakter `;` di `txbSelectedMacro` juga menghapus token macro di depannya; perubahan teks lain mengikuti perilaku editor biasa. Jika pencarian folder GMS gagal, nama masih dapat diketik manual. Ketersediaan project baru diperiksa saat **Process**.

### Process, Continue, dan Statistics

1. Pilih konfigurasi di `MacroQueue`, lalu tekan **Select**.
2. Tekan **Process**. Runner memeriksa seluruh antrean sebelum menjalankan langkah pertama.
3. Setelah form utama suatu macro ditutup dan langkahnya menjadi `OK`, tekan **Continue** untuk menjalankan macro berikutnya. Jika Behavior memakai `@cmdContinue;` di antara kedua blok, runner melanjutkannya otomatis.
4. Lihat hasil di `lbxStatistics` atau tekan **Copy Statistic** untuk menyalin seluruh statistik ke clipboard.

Status langkah adalah `PENDING`, `RUNNING`, `OK`, `ERROR`, dan `SKIPPED`. Jika sebuah langkah gagal, antrean berhenti dan langkah berikutnya menjadi `SKIPPED`. Durasi ditampilkan dalam detik dengan tiga angka desimal; waktu menunggu tombol **Continue** tidak dihitung. Waktu langkah berhenti saat form utama ditutup, termasuk waktu interaksi pengguna, sebelum token luar blok diproses. `MacroRunnerMenu` tidak dapat ditutup saat macro sedang berjalan, token luar blok sedang diproses, atau `MacroQueue` masih terbuka.

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

### Token di luar blok macro

Setelah sebuah blok `{...}` ditutup, MacroRunner dapat menjalankan token berikut. Token ini milik **runner**, bukan target di dalam form, dan diproses setelah form langkah tersebut ditutup.

| Token | Perilaku |
| --- | --- |
| `@ConvertToCurves;` | Menjalankan `ConvertToCurves` pada selection aktif. Jika tidak ada objek terpilih atau konversi gagal, catat warning dan lanjutkan. |
| `@ClearSelection;` | Melepas selection pada dokumen aktif. Kegagalan dicatat sebagai warning. |
| `@cmdContinue;` | Memulai macro berikutnya secara otomatis setelah token sebelumnya diproses. Hanya boleh berada di antara dua blok dan harus menjadi token terakhir sebelum blok berikutnya. |

Setiap token luar blok **wajib diakhiri `;`**. Token lain yang namanya valid secara sintaks tetapi belum dikenal dicatat sebagai warning. Token luar blok diproses sesuai urutan penulisan setelah callback penutupan form selesai; jika ingin mengonversi selection sebelum melepasnya, tulis `@ConvertToCurves;` sebelum `@ClearSelection;`.

Contoh antrean dua macro dalam `txbSelectedMacro`:

```text
AutoLabel; AutoSaveNCreate;
```

Dan `txbMacroBehavior` yang melanjutkan langkah kedua tanpa menekan **Continue**:

```text
{AutoLabel:AutoLabelWizard[txbLabel="/*~";@cmdSubmit]}
@ClearSelection;
@cmdContinue;
{AutoSaveNCreate:AutoSNC[@cmdProcess;@cmdClose]}
```

Contoh ini tetap memerlukan dokumen dan konfigurasi target yang sesuai. `@cmdContinue;` tidak boleh ditulis setelah blok terakhir. Jika token luar blok tidak dikenal atau gagal, langkah macro yang sudah selesai tetap `OK`; warning dirangkum di akhir antrean. Jika penjadwalan token gagal dan masih ada langkah berikutnya, lanjutkan secara manual melalui **Continue**. Error saat menjalankan instruksi di dalam blok macro tetap menghentikan antrean dengan status `ERROR` dan `SKIPPED` untuk langkah tersisa.

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

**Process** membaca ulang konfigurasi dan memeriksa semua project GMS. Untuk setiap blok Behavior eksplisit, runner memanggil `MRTargetBridge.ValidateBehavior` milik project target sebelum menjalankan langkah pertama. Jika validasi salah satu langkah gagal, antrean tidak mulai berjalan. Saat eksekusi, `MRTargetBridge.RunBehavior` target menjalankan instruksi; runner memerlukan callback konfirmasi dari bridge. Saat **Save**, runner memeriksa sintaks, posisi, dan pemisah token luar blok, tetapi belum memeriksa apakah nama tokennya dikenal. Token tersebut dijalankan oleh MacroRunner setelah form target ditutup. Pemeriksaan awal tidak menjamin kondisi dokumen, selection, atau file output tetap sama ketika langkah benar-benar berjalan.

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

Jika file JSON belum ada, runner membaca konfigurasi lama dari registry `RinCorelMacros/MacroRunner/SequencesV1`. Perubahan pertama melalui **Save**, **Remove**, **Clear**, **Up**, atau **Down** menulisnya ke JSON; registry lama tidak dihapus. Setelah JSON ada, file tersebut menjadi sumber konfigurasi. Konfigurasi lama tanpa `configId` bisa dimuat, tetapi **Save** berikutnya tetap memerlukan `ConfigID`. Urutan konfigurasi mengikuti susunan entri di JSON. Pilihan aktif dan statistik hanya berlaku selama sesi runner.

## Integrasi di CorelDRAW/VBE

Source di `src/classes/` adalah class module; pertahankan `(Name)` masing-masing saat memasangnya di project MacroRunner. File di `src/forms/` adalah code-behind untuk UserForm dengan nama yang sesuai, bukan file desain `.frm` lengkap. Siapkan kontrol yang disebut di kode form, termasuk `cmdContinue` pada `MacroRunnerMenu`, `txbConfigID` dan `txbMacroBehavior` pada `MacroSelection`, serta `frmMacroLists` (MSForms.Frame), `cmdUp`, dan `cmdDown` pada `MacroQueue`. Sertakan class `MRQueueRow` agar klik pada baris Queue memilih konfigurasi yang sesuai. Sertakan pula `MRBehaviorProgram.cls` dan standard module `MRDeferredStep.bas` untuk token luar blok.

MacroRunner mencari `.gms` di folder GMS milik profil CorelDRAW aktif melalui `Application.UserDataPath`. Project tujuan harus sudah dimuat oleh CorelDRAW. `MRCatalog` mempunyai registrasi form untuk `ExportRelated` dan `AutoDistributeUF` serta beberapa nama prosedur pembuka lain; keberadaan nama dalam daftar itu tidak otomatis memasang bridge atau menjamin kompatibilitas target.

Untuk alur manual, pasang `src/Integration/MRTargetBridge.bas` pada project GMS tujuan dan gabungkan `FormLifecycleSnippet.vba` ke UserForm utamanya. Form harus menyediakan `MRBindRunner`, `MRDetachRunner`, dan callback saat `UserForm_Terminate`, agar runner mengetahui kapan langkah selesai. Sesuaikan integrasi dengan lifecycle form tujuan; bridge generik menolak form yang sudah terbuka. `AutoDistributeUF` memakai bridge khusus dari project AutoDistributeUF.

Untuk Behavior eksplisit, target juga harus menyediakan implementasi `MRTargetBridge.ValidateBehavior` dan `MRTargetBridge.RunBehavior`, kontrak form/target/action, serta callback konfirmasi yang diharapkan `MRPresenter`. Template bridge generik di repo ini hanya menyediakan `OpenMacro`; integrasi Behavior untuk ExportRelated berada di project targetnya. Jika parser atau protokol Behavior diperbarui, perbarui salinan source bersama di MacroRunner dan target yang menggunakannya. Setelah pemasangan, jalankan **Debug > Compile** pada project terkait dan uji alur manual serta Behavior dengan dokumen contoh.

## Struktur source

| Lokasi | Peran |
| --- | --- |
| `src/forms/` | `MacroRunnerMenu`, `MacroQueue`, `MacroSelection` |
| `src/classes/MRPresenter.cls` | Koordinasi UI, preflight, eksekusi, dan callback |
| `src/classes/MRBehaviorParser.cls` | Parser grammar Behavior |
| `src/classes/MRBehaviorProgram.cls` | Memisahkan blok target dan token runner di luar blok |
| `src/classes/MRCatalog.cls` | Pencarian GMS dan resolusi project/form |
| `src/classes/MRRunModel.cls` | Status antrean dan statistik |
| `src/classes/MRSequenceStore.cls`, `MRConfigJson.cls` | Penyimpanan dan pembacaan konfigurasi |
| `src/classes/MRQueueRow.cls` | Event klik untuk judul dan urutan macro pada baris Queue |
| `src/modules/MRDeferredStep.bas` | Menjadwalkan token luar blok setelah callback penutupan form selesai |
| `src/Integration/` | Template bridge dan contoh hook lifecycle untuk target |

## Lisensi dan masukan

Project ini menggunakan lisensi MIT; lihat [LICENSE](LICENSE).

Source boleh dipelajari dan dikembangkan, dan issue/feedback tentang bug, edge case, CorelDRAW API, architecture, atau improvement sangat dihargai.
