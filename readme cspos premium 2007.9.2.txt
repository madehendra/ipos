CodeSuite Point Of Sales
Copyright 2007

Developer: made hendra
Contact: made.hendra@gmail.com
Website: http://balihosting.balidreamball.com
Phone: 081338414828


Major
 - Perbaikan dilakukan pada tahun
Minor
 - Perbaikan dilakukan pada bulan
Revision
 - Perubahan sistem
 - Perubahan pada struktur tabel
 - Perubahan tampilan (GUI)
 - Perubahan modul,form, component
 - Perubahan pada penyempurnaan baris kode tanpa adanya perubahan tampilan
 - Perbaikan bug pada modul tanpa mempengaruhi modul yang lain
 
Map
+ = Tambahan
- = Pengurangan
! = Update/Penyempurnaan 

Todo
! Ada konfirmasi apabila stock dalam kondisi minus, ketika dilakukan penjualan kasir atau penjualan non tunai
+ Ada semacam baloon peringatan apabila stock sudah mencapai kondisi min atau max pada systray

CPOS 2007.9.2
==========
! Bisa memasukkan barang berulang ulang (untuk barang yg sama) pada modul pembelian dan penjualan non tunai
! Bug, pada laporan laba rugi, nilai pembelian yg dihitung tertera nilai total nya. Padahal seharusnya yg dihitung adalah nilai subtotalnya. Jadinya diskon jadi sangat besar.

CPOS 1.9.1
==========
Release : 10:23 AM Saturday, September 01, 2007
! Window setfocus yg lebih baik.
! Ada pilihan check box untuk mempercepat melakukan pelunasan (tunai) pada modul pelunasan hutang/piutang
! Tambahan baris "Print by Username" pada setiap cetakan kasir/penjualan non tunai dalam bentuk struk. 

CPOS 1.8.4
==========
Release : 11:15 AM Wednesday, August 22, 2007

!PENTING
Ada sedikit perubahan pada library.
Setelah cspos v1.8.4 di update disarankan kepada seluruh customer untuk melakukan
seting ulang pada modul General Settings...

Lakukan Update database terlebih dahulu sebelum program dijalankan.

!UPDATE
+ Penambahan cetakan barcode, Task to DO >> Kasir >> Cetak Stiker Barcode
+ Laporan pemasukan
+ Laporan pengeluaran
+ Laporan penjualan Group By Member
! Kolom Qty pada Penjualan Kasir dapat di set otomatis nilainya. Apakah otomatis 1, 2, dsb
  Modul General Settings...
! Kolom Discount % pada modul pembelian dapat di set otomatis nilainya  
  Modul General Settings...
+ Tambahan modul kategory inventory dan non inventory
! Beberapa laporan seperti laporan kartu piutang, hutang, dan pelunasan piutang/hutang mengalami sedikit revisi.Bug tgl, sering tampil terbalik.

CPOS 1.1.3
==========
! Penyempurnaan tampilan grid inventory
! Penyempurnaan tampilan grid member
! Penyempurnaan tampilan grid supplier
! Warna background untuk modul penjualan, retur penjualan, dan pelunasan piutang dirubah menjadi warna putih untuk membedakan dengan modul pembelian 

CPOS 1.1.2
==========
+ Kasir bisa mencetak transaksi yg sudah terjadi sebelumnya lewat menu Transaksi > Cetak atau Pembatalan Penjualan Kasir
+ Tambahan menu About

CPOS 1.1.1
==========
! Penyempurnaan laporan penjualan kasir
! Penyempurnaan laporan pembatalan penjualan kasir
+ Tambahan modul/form penjualan harian kasir
! Pada transaksi penjualan kasir dan penjualan non tunai, harga bisa dirubah tetapi harga tidak boleh kurang dari harga beli.
+ Tambahan modul General Settings
- Modul Setup printer kasir dihilangkan, dijadikan satu pada modul General Settings
+ Cetakan untuk penjualan non tunai bisa berbentuk struk

CPOS 1.1.0
==========
+ Struktur table untuk table pemasukan dan pengeluaran, serta posnya
! Tersedianya Plafond untuk customer
+ Transaksi Master pos pemasukan dan pengeluaran
+ Transaksi Pemasukan dan Pengeluaran
! Proses penyimpanan pada Form Stockopname tidak sempurna
! Penyempurnaan tampilan dan proses input pada master : supplier, inventory, customer, golongan stock, satuan.
+ Laporan Laba Rugi
! Setiap kali menginput qty selain 1 pada form Penjualan Kasir, jumlah yg dihitung tetap 1.