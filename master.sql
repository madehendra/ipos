/*
Navicat MySQL Data Transfer

Source Server         : localhost
Source Server Version : 50505
Source Host           : 127.0.0.1:3306
Source Database       : sw

Target Server Type    : MYSQL
Target Server Version : 50505
File Encoding         : 65001

Date: 2018-12-19 12:26:26
*/

SET FOREIGN_KEY_CHECKS=0;

-- ----------------------------
-- Table structure for akun
-- ----------------------------
DROP TABLE IF EXISTS `akun`;
CREATE TABLE `akun` (
  `kodeakun` char(20) NOT NULL,
  `keterangan` char(200) DEFAULT NULL,
  `jenis` char(20) DEFAULT NULL,
  `budget` double DEFAULT NULL,
  PRIMARY KEY (`kodeakun`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for akunbiaya
-- ----------------------------
DROP TABLE IF EXISTS `akunbiaya`;
CREATE TABLE `akunbiaya` (
  `kodeakun` char(20) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for akunkas
-- ----------------------------
DROP TABLE IF EXISTS `akunkas`;
CREATE TABLE `akunkas` (
  `kodeakun` char(20) NOT NULL,
  `username` char(20) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for anggota
-- ----------------------------
DROP TABLE IF EXISTS `anggota`;
CREATE TABLE `anggota` (
  `kodeanggota` char(20) NOT NULL,
  `kodedep` char(20) NOT NULL,
  `nama` char(40) DEFAULT NULL,
  `alamat` char(50) DEFAULT NULL,
  `plafond` double DEFAULT NULL,
  `status` char(1) DEFAULT NULL,
  `tgl` date DEFAULT NULL,
  `nopeg` varchar(20) DEFAULT NULL,
  `telp` varchar(20) DEFAULT NULL,
  `kodeupline` varchar(20) DEFAULT NULL,
  `nlevel` double DEFAULT NULL,
  `dd` double DEFAULT NULL,
  `diskon` double DEFAULT NULL,
  `lastactivity` date DEFAULT NULL,
  PRIMARY KEY (`kodeanggota`),
  KEY `idxAnggota` (`kodeanggota`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for batalkasir
-- ----------------------------
DROP TABLE IF EXISTS `batalkasir`;
CREATE TABLE `batalkasir` (
  `nomorkasir` varchar(30) DEFAULT NULL,
  `kodestock` char(100) DEFAULT NULL,
  `qty` double DEFAULT NULL,
  `harga` double DEFAULT NULL,
  `jumlah` double DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for bg
-- ----------------------------
DROP TABLE IF EXISTS `bg`;
CREATE TABLE `bg` (
  `tableid` bigint(20) NOT NULL AUTO_INCREMENT,
  `nomorpelunasanpiutang` varchar(20) DEFAULT NULL,
  `reff` varchar(20) DEFAULT NULL,
  `jumlah` double DEFAULT NULL,
  `jatuhtempo` date DEFAULT NULL,
  PRIMARY KEY (`tableid`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for biaya
-- ----------------------------
DROP TABLE IF EXISTS `biaya`;
CREATE TABLE `biaya` (
  `nomorbiaya` char(20) NOT NULL,
  `kodeakun` char(20) NOT NULL,
  `keterangan` char(200) DEFAULT NULL,
  `jumlah` char(20) DEFAULT NULL,
  `budget` double DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for bukubesar
-- ----------------------------
DROP TABLE IF EXISTS `bukubesar`;
CREATE TABLE `bukubesar` (
  `idbukubesar` bigint(20) NOT NULL AUTO_INCREMENT,
  `kodeakun` char(20) NOT NULL,
  `username` char(20) NOT NULL,
  `tgl` date DEFAULT NULL,
  `faktur` char(20) DEFAULT NULL,
  `status` char(20) DEFAULT NULL,
  `keterangan` char(200) DEFAULT NULL,
  `debet` double DEFAULT NULL,
  `kredit` double DEFAULT NULL,
  `kas` char(20) DEFAULT NULL,
  `datetime` datetime DEFAULT NULL,
  `kodecostcenter` char(20) NOT NULL,
  `kodestock` varchar(255) DEFAULT NULL,
  PRIMARY KEY (`idbukubesar`),
  KEY `idxDateTime` (`datetime`),
  KEY `idxFaktur` (`faktur`),
  KEY `idxKodeAkun` (`kodeakun`)
) ENGINE=InnoDB AUTO_INCREMENT=14438 DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for buyback
-- ----------------------------
DROP TABLE IF EXISTS `buyback`;
CREATE TABLE `buyback` (
  `nomorbuyback` char(20) NOT NULL DEFAULT '',
  `kodegudang` char(20) NOT NULL DEFAULT '',
  `tgl` date DEFAULT NULL,
  `kodestock` char(30) DEFAULT NULL,
  `qty` double DEFAULT NULL,
  `harga` double DEFAULT NULL,
  `kodesatuan` char(20) DEFAULT NULL,
  `discount` double DEFAULT NULL,
  `jumlah` double DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for config
-- ----------------------------
DROP TABLE IF EXISTS `config`;
CREATE TABLE `config` (
  `jenis` char(2) NOT NULL,
  `keterangan` varchar(255) DEFAULT NULL,
  `tipe` char(1) DEFAULT NULL,
  `label` varchar(255) DEFAULT NULL,
  `modul` varchar(255) DEFAULT NULL,
  PRIMARY KEY (`jenis`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for costcenter
-- ----------------------------
DROP TABLE IF EXISTS `costcenter`;
CREATE TABLE `costcenter` (
  `kodecostcenter` char(20) NOT NULL,
  `keterangan` char(20) DEFAULT NULL,
  PRIMARY KEY (`kodecostcenter`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for dbver
-- ----------------------------
DROP TABLE IF EXISTS `dbver`;
CREATE TABLE `dbver` (
  `ver` varchar(255) DEFAULT NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for dep
-- ----------------------------
DROP TABLE IF EXISTS `dep`;
CREATE TABLE `dep` (
  `kodedep` char(20) NOT NULL,
  `keterangan` char(50) DEFAULT NULL,
  PRIMARY KEY (`kodedep`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for estimasi
-- ----------------------------
DROP TABLE IF EXISTS `estimasi`;
CREATE TABLE `estimasi` (
  `nomorpenjualan` char(20) NOT NULL,
  `kodegudang` char(20) NOT NULL,
  `tgl` date DEFAULT NULL,
  `kodestock` char(30) DEFAULT NULL,
  `qty` double DEFAULT NULL,
  `harga` double DEFAULT NULL,
  `kodesatuan` char(20) DEFAULT NULL,
  `discount` double DEFAULT NULL,
  `jumlah` double DEFAULT NULL,
  `hb` double DEFAULT NULL,
  `tunai` double DEFAULT NULL,
  `piutang` double DEFAULT NULL,
  `statuslunas` char(1) DEFAULT '0'
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for formlevel
-- ----------------------------
DROP TABLE IF EXISTS `formlevel`;
CREATE TABLE `formlevel` (
  `nama` char(50) NOT NULL,
  `userLevel` char(3) NOT NULL,
  `status` char(10) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for golongan
-- ----------------------------
DROP TABLE IF EXISTS `golongan`;
CREATE TABLE `golongan` (
  `kodegolongan` char(11) NOT NULL,
  `keterangan` char(40) DEFAULT NULL,
  PRIMARY KEY (`kodegolongan`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for gudang
-- ----------------------------
DROP TABLE IF EXISTS `gudang`;
CREATE TABLE `gudang` (
  `kodegudang` char(20) NOT NULL,
  `keterangan` char(20) DEFAULT NULL,
  `lstatus` char(1) DEFAULT NULL,
  PRIMARY KEY (`kodegudang`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for infobc
-- ----------------------------
DROP TABLE IF EXISTS `infobc`;
CREATE TABLE `infobc` (
  `kodestock` varchar(20) NOT NULL DEFAULT '',
  `keterangan` text,
  `datetime` datetime DEFAULT NULL,
  `username` varchar(255) DEFAULT NULL,
  `flagdel` varchar(1) DEFAULT '1',
  PRIMARY KEY (`kodestock`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for jurnalumum
-- ----------------------------
DROP TABLE IF EXISTS `jurnalumum`;
CREATE TABLE `jurnalumum` (
  `kodeakun` char(20) NOT NULL,
  `nomorjurnalumum` char(20) NOT NULL,
  `tgl` date DEFAULT NULL,
  `keterangan` char(200) DEFAULT NULL,
  `debet` double DEFAULT NULL,
  `kredit` double DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for kartuhutang
-- ----------------------------
DROP TABLE IF EXISTS `kartuhutang`;
CREATE TABLE `kartuhutang` (
  `kodesupplier` char(6) NOT NULL,
  `username` char(20) NOT NULL,
  `id` bigint(20) NOT NULL AUTO_INCREMENT,
  `status` char(2) DEFAULT NULL,
  `nomorkartuhutang` varchar(30) DEFAULT NULL,
  `tgl` date DEFAULT NULL,
  `keterangan` varchar(255) DEFAULT NULL,
  `debet` double DEFAULT NULL,
  `kredit` double DEFAULT NULL,
  `datetime` datetime DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=33 DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for kartupiutang
-- ----------------------------
DROP TABLE IF EXISTS `kartupiutang`;
CREATE TABLE `kartupiutang` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT,
  `username` char(20) NOT NULL,
  `kodeanggota` char(20) NOT NULL,
  `status` char(2) DEFAULT NULL,
  `nomorkartupiutang` varchar(30) DEFAULT NULL,
  `tgl` date DEFAULT NULL,
  `keterangan` varchar(255) DEFAULT NULL,
  `debet` double DEFAULT NULL,
  `kredit` double DEFAULT NULL,
  `datetime` datetime DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=8 DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for kartustock
-- ----------------------------
DROP TABLE IF EXISTS `kartustock`;
CREATE TABLE `kartustock` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT,
  `kodegudang` char(20) NOT NULL,
  `username` char(20) NOT NULL,
  `status` char(2) DEFAULT NULL,
  `nomor` varchar(20) DEFAULT NULL,
  `tgl` date DEFAULT NULL,
  `kodestock` varchar(20) DEFAULT NULL,
  `qty` double DEFAULT NULL,
  `debet` double DEFAULT NULL,
  `kredit` double DEFAULT NULL,
  `harga` double DEFAULT '0',
  `hp` double DEFAULT NULL,
  `keterangan` varchar(200) DEFAULT NULL,
  `datetime` datetime DEFAULT '0000-00-00 00:00:00',
  `tothp` double DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=4933 DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for kasir
-- ----------------------------
DROP TABLE IF EXISTS `kasir`;
CREATE TABLE `kasir` (
  `nomorkasir` varchar(30) DEFAULT NULL,
  `kodestock` char(100) DEFAULT NULL,
  `qty` double DEFAULT NULL,
  `harga` double DEFAULT NULL,
  `jumlah` double DEFAULT NULL,
  `hargabeli` double DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for keyapp
-- ----------------------------
DROP TABLE IF EXISTS `keyapp`;
CREATE TABLE `keyapp` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `tokenapp` varchar(255) DEFAULT NULL,
  `keyapp` varchar(255) DEFAULT NULL,
  `tgl` date DEFAULT NULL,
  `tglbb` date DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=2 DEFAULT CHARSET=latin1 ROW_FORMAT=COMPACT;

-- ----------------------------
-- Table structure for kontrakstock
-- ----------------------------
DROP TABLE IF EXISTS `kontrakstock`;
CREATE TABLE `kontrakstock` (
  `tgl` date DEFAULT NULL,
  `kodeanggota` varchar(20) DEFAULT NULL,
  `kodestock` varchar(20) DEFAULT NULL,
  `hargajual` double DEFAULT NULL,
  `hargakontrak` double DEFAULT NULL,
  `username` varchar(40) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for member
-- ----------------------------
DROP TABLE IF EXISTS `member`;
CREATE TABLE `member` (
  `nama` varchar(255) DEFAULT NULL,
  `nohp` varchar(255) DEFAULT NULL,
  `ip` double(255,3) DEFAULT NULL,
  `group` varchar(255) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for memberbalance
-- ----------------------------
DROP TABLE IF EXISTS `memberbalance`;
CREATE TABLE `memberbalance` (
  `memberbalanceid` char(20) NOT NULL,
  `kodeanggota` char(20) DEFAULT NULL,
  `balance` double DEFAULT NULL,
  PRIMARY KEY (`memberbalanceid`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for memberorder
-- ----------------------------
DROP TABLE IF EXISTS `memberorder`;
CREATE TABLE `memberorder` (
  `nomormemberorder` char(20) NOT NULL,
  `kodegudang` char(20) NOT NULL,
  `tgl` date DEFAULT NULL,
  `kodestock` char(30) DEFAULT NULL,
  `qty` double DEFAULT NULL,
  `harga` double DEFAULT NULL,
  `kodesatuan` char(20) DEFAULT NULL,
  `discount` double DEFAULT NULL,
  `jumlah` double DEFAULT NULL,
  `hb` double DEFAULT NULL,
  `tunai` double DEFAULT NULL,
  `piutang` double DEFAULT NULL,
  `statuslunas` char(1) DEFAULT '0',
  `nourut` double DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for membertopup
-- ----------------------------
DROP TABLE IF EXISTS `membertopup`;
CREATE TABLE `membertopup` (
  `nomormembertopup` char(20) NOT NULL,
  `tgl` date DEFAULT NULL,
  `kodeanggota` varchar(255) DEFAULT NULL,
  `keterangan` char(200) DEFAULT NULL,
  `debet` double DEFAULT '0',
  `kredit` double DEFAULT '0',
  `lstatus` char(20) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for menu
-- ----------------------------
DROP TABLE IF EXISTS `menu`;
CREATE TABLE `menu` (
  `kode` varchar(10) DEFAULT NULL,
  `menulevel` char(3) DEFAULT NULL,
  `filename` varchar(255) DEFAULT NULL,
  `urut` smallint(6) DEFAULT '0',
  `status` char(1) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for msthadiah
-- ----------------------------
DROP TABLE IF EXISTS `msthadiah`;
CREATE TABLE `msthadiah` (
  `kodehadiah` varchar(50) NOT NULL,
  `keterangan` varchar(255) DEFAULT NULL,
  `poin` double DEFAULT NULL,
  `status` char(1) DEFAULT '1',
  `gambar` mediumblob,
  PRIMARY KEY (`kodehadiah`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for mutasikasbank
-- ----------------------------
DROP TABLE IF EXISTS `mutasikasbank`;
CREATE TABLE `mutasikasbank` (
  `nomormutasikasbank` varchar(20) NOT NULL,
  `dariakun` varchar(20) NOT NULL,
  `keakun` varchar(20) NOT NULL,
  `total` double DEFAULT NULL,
  `tgl` date DEFAULT NULL,
  `datetime` datetime DEFAULT NULL,
  `username` varchar(50) DEFAULT NULL,
  `debet` double DEFAULT NULL,
  `kredit` double DEFAULT NULL,
  `kodecostcenter` varchar(20) DEFAULT NULL,
  `keterangan` varchar(200) DEFAULT NULL,
  PRIMARY KEY (`nomormutasikasbank`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for mutasistock
-- ----------------------------
DROP TABLE IF EXISTS `mutasistock`;
CREATE TABLE `mutasistock` (
  `nomormutasistock` char(20) NOT NULL,
  `kodestock` bigint(20) NOT NULL,
  `gudangdari` char(20) DEFAULT NULL,
  `gudangke` char(20) DEFAULT NULL,
  `qty` double DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for nomorfaktur
-- ----------------------------
DROP TABLE IF EXISTS `nomorfaktur`;
CREATE TABLE `nomorfaktur` (
  `nomorfaktur` varchar(255) NOT NULL,
  `modul` varchar(255) NOT NULL,
  PRIMARY KEY (`nomorfaktur`,`modul`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for op_tbl
-- ----------------------------
DROP TABLE IF EXISTS `op_tbl`;
CREATE TABLE `op_tbl` (
  `op_frefix` varchar(255) DEFAULT NULL,
  `op_name` varchar(255) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for ord
-- ----------------------------
DROP TABLE IF EXISTS `ord`;
CREATE TABLE `ord` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT,
  `kodestock` varchar(255) DEFAULT NULL,
  `barcode` varchar(255) DEFAULT NULL,
  `tgl` date DEFAULT NULL,
  `qty` smallint(6) DEFAULT NULL,
  `kodemember` varchar(255) DEFAULT NULL,
  `lstatus` char(255) DEFAULT 'N',
  PRIMARY KEY (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for orderan
-- ----------------------------
DROP TABLE IF EXISTS `orderan`;
CREATE TABLE `orderan` (
  `orderid` int(11) NOT NULL AUTO_INCREMENT,
  `reffid` varchar(30) DEFAULT '',
  `status` char(1) DEFAULT '',
  `tgl` date DEFAULT NULL,
  `kodeanggota` varchar(30) DEFAULT NULL,
  `kodestock` int(11) DEFAULT NULL,
  `barcode` varchar(50) DEFAULT NULL,
  `debet` double DEFAULT '0',
  `kredit` double DEFAULT '0',
  `keterangan` text,
  `username` varchar(50) DEFAULT NULL,
  `datetime` datetime DEFAULT NULL,
  PRIMARY KEY (`orderid`),
  KEY `kodeanggota` (`kodeanggota`),
  CONSTRAINT `orderan_ibfk_1` FOREIGN KEY (`kodeanggota`) REFERENCES `anggota` (`kodeanggota`) ON UPDATE CASCADE
) ENGINE=InnoDB AUTO_INCREMENT=3426 DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for packing
-- ----------------------------
DROP TABLE IF EXISTS `packing`;
CREATE TABLE `packing` (
  `nopacking` varchar(20) DEFAULT NULL,
  `kodestock` varchar(20) DEFAULT NULL,
  `jumlah` double DEFAULT NULL,
  `status` char(2) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for pelunasanhutang
-- ----------------------------
DROP TABLE IF EXISTS `pelunasanhutang`;
CREATE TABLE `pelunasanhutang` (
  `nomorpelunasanhutang` char(20) NOT NULL,
  `nomorpembelian` char(20) NOT NULL,
  `id` bigint(20) unsigned NOT NULL AUTO_INCREMENT,
  `hutang` double DEFAULT NULL,
  `discount` double DEFAULT NULL,
  `pelunasan` double DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=13 DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for pelunasanpiutang
-- ----------------------------
DROP TABLE IF EXISTS `pelunasanpiutang`;
CREATE TABLE `pelunasanpiutang` (
  `nomorpenjualan` char(20) NOT NULL,
  `nomorpelunasanpiutang` char(20) NOT NULL,
  `piutang` double DEFAULT NULL,
  `discount` double DEFAULT NULL,
  `pelunasan` double DEFAULT NULL,
  KEY `idxPelunasanpiutang` (`nomorpelunasanpiutang`,`nomorpenjualan`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for pembelian
-- ----------------------------
DROP TABLE IF EXISTS `pembelian`;
CREATE TABLE `pembelian` (
  `nomorpembelian` char(20) NOT NULL,
  `kodegudang` char(20) NOT NULL,
  `tgl` date DEFAULT NULL,
  `kodestock` char(30) DEFAULT NULL,
  `qty` double DEFAULT NULL,
  `harga` double DEFAULT NULL,
  `kodesatuan` char(20) DEFAULT NULL,
  `discount` double DEFAULT NULL,
  `jumlah` double DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for pencairanbg
-- ----------------------------
DROP TABLE IF EXISTS `pencairanbg`;
CREATE TABLE `pencairanbg` (
  `tableid` varchar(20) NOT NULL DEFAULT '',
  `nomorpelunasanpiutang` varchar(20) DEFAULT NULL,
  `kodeakun` varchar(20) DEFAULT NULL,
  `username` varchar(20) DEFAULT NULL,
  `date` date DEFAULT NULL,
  `datetime` datetime DEFAULT NULL,
  `keterangan` varchar(200) DEFAULT NULL,
  PRIMARY KEY (`tableid`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for pendingtrans
-- ----------------------------
DROP TABLE IF EXISTS `pendingtrans`;
CREATE TABLE `pendingtrans` (
  `nomorpenjualan` char(20) NOT NULL,
  `kodegudang` char(20) DEFAULT NULL,
  `tgl` date DEFAULT NULL,
  `kodestock` char(30) DEFAULT NULL,
  `qty` double DEFAULT NULL,
  `harga` double DEFAULT NULL,
  `kodesatuan` char(20) DEFAULT NULL,
  `discount` double DEFAULT NULL,
  `jumlah` double DEFAULT NULL,
  `hb` double DEFAULT NULL,
  `tunai` double DEFAULT NULL,
  `piutang` double DEFAULT NULL,
  `statuslunas` char(1) DEFAULT '0',
  `urutfaktur` double DEFAULT NULL,
  `bv` double NOT NULL,
  `userid` varchar(255) DEFAULT NULL,
  KEY `idxPenjualan` (`nomorpenjualan`) USING BTREE
) ENGINE=InnoDB DEFAULT CHARSET=latin1 ROW_FORMAT=COMPACT;

-- ----------------------------
-- Table structure for penjualan
-- ----------------------------
DROP TABLE IF EXISTS `penjualan`;
CREATE TABLE `penjualan` (
  `nomorpenjualan` char(20) NOT NULL,
  `kodegudang` char(20) DEFAULT NULL,
  `tgl` date DEFAULT NULL,
  `kodestock` char(30) DEFAULT NULL,
  `qty` double DEFAULT NULL,
  `harga` double DEFAULT NULL,
  `kodesatuan` char(20) DEFAULT NULL,
  `discount` double DEFAULT NULL,
  `jumlah` double DEFAULT NULL,
  `hb` double DEFAULT NULL,
  `tunai` double DEFAULT NULL,
  `piutang` double DEFAULT NULL,
  `statuslunas` char(1) DEFAULT '0',
  `urutfaktur` double DEFAULT NULL,
  `bv` double NOT NULL,
  KEY `idxPenjualan` (`nomorpenjualan`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for periode
-- ----------------------------
DROP TABLE IF EXISTS `periode`;
CREATE TABLE `periode` (
  `Status` char(1) NOT NULL DEFAULT '0',
  `Kode` char(4) NOT NULL DEFAULT '',
  `Awal` date DEFAULT NULL,
  `Akhir` date DEFAULT NULL,
  `Keterangan` char(50) DEFAULT NULL,
  PRIMARY KEY (`Status`,`Kode`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for persediaan
-- ----------------------------
DROP TABLE IF EXISTS `persediaan`;
CREATE TABLE `persediaan` (
  `kodestock` varchar(20) NOT NULL,
  `harga` double DEFAULT NULL,
  `qty` double DEFAULT NULL,
  `persediaan` double DEFAULT NULL,
  PRIMARY KEY (`kodestock`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for po
-- ----------------------------
DROP TABLE IF EXISTS `po`;
CREATE TABLE `po` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT,
  `statusorder` char(1) DEFAULT '0',
  `tgl` date DEFAULT NULL,
  `kodeso` varchar(20) DEFAULT NULL,
  `kodestock` varchar(20) DEFAULT NULL,
  `qty` double DEFAULT NULL,
  `harga` double DEFAULT NULL,
  `diskonpenjualan` double DEFAULT NULL,
  `statuspembelian` char(1) DEFAULT '0',
  `statuscancel` char(1) DEFAULT '0',
  `fakturpembelian` varchar(20) DEFAULT NULL,
  `datetime` datetime DEFAULT NULL,
  `username` varchar(20) DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for poinhadiah
-- ----------------------------
DROP TABLE IF EXISTS `poinhadiah`;
CREATE TABLE `poinhadiah` (
  `faktur` varchar(255) DEFAULT NULL,
  `tgl` date DEFAULT NULL,
  `kodeanggota` varchar(255) DEFAULT NULL,
  `poinhadiah` int(11) DEFAULT NULL,
  `tukar` int(11) DEFAULT NULL,
  `tukardate` date DEFAULT NULL,
  `exdate` date DEFAULT NULL,
  `status` char(1) DEFAULT NULL,
  `keterangan` varchar(255) DEFAULT ''
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for pos
-- ----------------------------
DROP TABLE IF EXISTS `pos`;
CREATE TABLE `pos` (
  `kodepos` char(30) NOT NULL,
  `keterangan` char(100) DEFAULT NULL,
  `jenis` char(1) DEFAULT NULL,
  PRIMARY KEY (`kodepos`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for prive
-- ----------------------------
DROP TABLE IF EXISTS `prive`;
CREATE TABLE `prive` (
  `nomorprive` varchar(20) NOT NULL,
  `akunkas` varchar(20) NOT NULL,
  `akunprive` varchar(20) NOT NULL,
  `total` double DEFAULT NULL,
  `tgl` date DEFAULT NULL,
  `datetime` datetime DEFAULT NULL,
  `username` varchar(50) DEFAULT NULL,
  `kodecostcenter` varchar(20) DEFAULT NULL,
  `keterangan` varchar(200) DEFAULT NULL,
  PRIMARY KEY (`nomorprive`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for returpembelian
-- ----------------------------
DROP TABLE IF EXISTS `returpembelian`;
CREATE TABLE `returpembelian` (
  `kodesatuan` char(20) NOT NULL,
  `kodestock` bigint(20) NOT NULL,
  `nomorreturpembelian` char(20) NOT NULL,
  `tgl` date DEFAULT NULL,
  `qty` double DEFAULT NULL,
  `harga` double DEFAULT NULL,
  `discount` double DEFAULT NULL,
  `jumlah` double DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for returpenjualan
-- ----------------------------
DROP TABLE IF EXISTS `returpenjualan`;
CREATE TABLE `returpenjualan` (
  `kodesatuan` char(20) NOT NULL,
  `kodestock` bigint(20) NOT NULL,
  `nomorreturpenjualan` char(20) NOT NULL,
  `tgl` date DEFAULT NULL,
  `qty` double DEFAULT NULL,
  `harga` double DEFAULT NULL,
  `discount` double DEFAULT NULL,
  `jumlah` double DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for salesman
-- ----------------------------
DROP TABLE IF EXISTS `salesman`;
CREATE TABLE `salesman` (
  `kodesalesman` varchar(20) NOT NULL,
  `nama` varchar(100) DEFAULT NULL,
  `alamat` varchar(100) DEFAULT NULL,
  `telp` varchar(50) DEFAULT NULL,
  PRIMARY KEY (`kodesalesman`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for satuan
-- ----------------------------
DROP TABLE IF EXISTS `satuan`;
CREATE TABLE `satuan` (
  `kodesatuan` char(20) NOT NULL,
  `keterangan` char(40) DEFAULT NULL,
  PRIMARY KEY (`kodesatuan`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for stock
-- ----------------------------
DROP TABLE IF EXISTS `stock`;
CREATE TABLE `stock` (
  `kodesatuan` char(20) DEFAULT NULL,
  `kodegolongan` char(11) NOT NULL,
  `kodestock` bigint(20) NOT NULL AUTO_INCREMENT,
  `barcode` varchar(20) DEFAULT NULL,
  `nama` varchar(40) DEFAULT NULL,
  `hargabeli` double DEFAULT NULL,
  `hargajual` double DEFAULT NULL,
  `cogs` double DEFAULT NULL,
  `jenis` varchar(2) NOT NULL DEFAULT '1',
  `outsource` varchar(2) CHARACTER SET latin1 COLLATE latin1_bin DEFAULT 'T',
  `kodesupplier` varchar(20) DEFAULT NULL,
  `asbiaya` char(1) DEFAULT NULL,
  `saldostock` double DEFAULT NULL,
  `poin` double DEFAULT NULL,
  `diskonpenjualan` double DEFAULT NULL,
  `statusnonaktif` smallint(6) DEFAULT '0',
  `datetime` date DEFAULT NULL,
  `bv` double(255,0) NOT NULL,
  `tcogs` double(255,0) DEFAULT NULL,
  `stok` double(255,0) NOT NULL DEFAULT '0',
  PRIMARY KEY (`kodestock`)
) ENGINE=InnoDB AUTO_INCREMENT=107048 DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for stockopname
-- ----------------------------
DROP TABLE IF EXISTS `stockopname`;
CREATE TABLE `stockopname` (
  `nomorstockopname` char(20) NOT NULL,
  `kodestock` bigint(20) NOT NULL,
  `kodegudang` char(20) NOT NULL,
  `adjust` double DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for stok
-- ----------------------------
DROP TABLE IF EXISTS `stok`;
CREATE TABLE `stok` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT,
  `kodegudang` char(20) NOT NULL,
  `username` char(20) NOT NULL,
  `status` char(2) DEFAULT NULL,
  `nomor` varchar(20) DEFAULT NULL,
  `tgl` date DEFAULT NULL,
  `kodestock` varchar(20) DEFAULT NULL,
  `qty` double DEFAULT NULL,
  `debet` double DEFAULT NULL,
  `kredit` double DEFAULT NULL,
  `harga` double DEFAULT '0',
  `hp` double DEFAULT NULL,
  `keterangan` varchar(200) DEFAULT NULL,
  `datetime` datetime DEFAULT '0000-00-00 00:00:00',
  `tothp` double DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 ROW_FORMAT=COMPACT;

-- ----------------------------
-- Table structure for supplier
-- ----------------------------
DROP TABLE IF EXISTS `supplier`;
CREATE TABLE `supplier` (
  `kodesupplier` char(6) NOT NULL,
  `kodeakun` char(20) NOT NULL DEFAULT '',
  `nama` char(40) DEFAULT NULL,
  `alamat` char(50) DEFAULT NULL,
  `telepon` char(30) DEFAULT NULL,
  `fax` char(30) DEFAULT NULL,
  `kota` char(20) DEFAULT NULL,
  PRIMARY KEY (`kodesupplier`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for supplierbalance
-- ----------------------------
DROP TABLE IF EXISTS `supplierbalance`;
CREATE TABLE `supplierbalance` (
  `supplierbalanceid` char(20) NOT NULL DEFAULT '',
  `kodesupplier` char(20) DEFAULT NULL,
  `balance` double DEFAULT NULL,
  PRIMARY KEY (`supplierbalanceid`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for totbatalkasir
-- ----------------------------
DROP TABLE IF EXISTS `totbatalkasir`;
CREATE TABLE `totbatalkasir` (
  `nomorkasir` varchar(30) DEFAULT NULL,
  `subtotal` double DEFAULT NULL,
  `discount` double DEFAULT NULL,
  `total` double DEFAULT NULL,
  `tunai` double DEFAULT NULL,
  `username` varchar(100) DEFAULT NULL,
  `tgl` date DEFAULT NULL,
  `datetime` datetime DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for totbiaya
-- ----------------------------
DROP TABLE IF EXISTS `totbiaya`;
CREATE TABLE `totbiaya` (
  `nomorbiaya` char(20) NOT NULL,
  `username` char(20) NOT NULL,
  `kodeakun` char(20) NOT NULL,
  `kodecostcenter` char(20) NOT NULL,
  `jumlah` double DEFAULT NULL,
  `tgl` date DEFAULT NULL,
  `datetime` datetime DEFAULT NULL,
  PRIMARY KEY (`nomorbiaya`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for totbuyback
-- ----------------------------
DROP TABLE IF EXISTS `totbuyback`;
CREATE TABLE `totbuyback` (
  `nomorbuyback` char(20) NOT NULL DEFAULT '',
  `kodecostcenter` char(20) NOT NULL DEFAULT '',
  `kodeakun` char(20) NOT NULL DEFAULT '',
  `kodeanggota` char(20) NOT NULL DEFAULT '',
  `username` char(20) NOT NULL DEFAULT '',
  `fakturasli` char(20) DEFAULT NULL,
  `tgl` date DEFAULT NULL,
  `jthtmp` date DEFAULT NULL,
  `ppn` double DEFAULT NULL,
  `persdisc` double DEFAULT NULL,
  `persdisc2` double DEFAULT NULL,
  `subtotal` double DEFAULT NULL,
  `pajak` double DEFAULT NULL,
  `discount` double DEFAULT NULL,
  `discount2` double DEFAULT NULL,
  `total` double DEFAULT NULL,
  `tunai` double DEFAULT NULL,
  `hutang` double DEFAULT NULL,
  `datetime` datetime DEFAULT NULL,
  `flaglunas` varchar(1) DEFAULT NULL,
  `kodesalesman` varchar(20) DEFAULT NULL,
  `statusbuyback` char(1) DEFAULT '0',
  `kodegudang` varchar(20) DEFAULT NULL,
  PRIMARY KEY (`nomorbuyback`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for totestimasi
-- ----------------------------
DROP TABLE IF EXISTS `totestimasi`;
CREATE TABLE `totestimasi` (
  `nomorpenjualan` char(20) NOT NULL,
  `kodeanggota` char(6) NOT NULL,
  `username` char(20) NOT NULL,
  `kodeakun` char(20) NOT NULL,
  `kodecostcenter` char(20) NOT NULL,
  `fakturasli` char(20) DEFAULT NULL,
  `tgl` date DEFAULT NULL,
  `jthtmp` date DEFAULT NULL,
  `ppn` double DEFAULT NULL,
  `persdisc` double DEFAULT NULL,
  `persdisc2` double DEFAULT NULL,
  `subtotal` double DEFAULT NULL,
  `pajak` double DEFAULT NULL,
  `discount` double DEFAULT NULL,
  `discount2` double DEFAULT NULL,
  `total` double DEFAULT NULL,
  `tunai` double DEFAULT NULL,
  `piutang` double DEFAULT NULL,
  `datetime` datetime DEFAULT NULL,
  `flaglunas` varchar(1) DEFAULT NULL,
  `kodesalesman` varchar(20) DEFAULT NULL,
  `komisi` double DEFAULT NULL,
  `totalbiaya` double DEFAULT NULL,
  PRIMARY KEY (`nomorpenjualan`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for totjurnalumum
-- ----------------------------
DROP TABLE IF EXISTS `totjurnalumum`;
CREATE TABLE `totjurnalumum` (
  `nomorjurnalumum` char(20) NOT NULL,
  `username` char(20) NOT NULL,
  `kodecostcenter` char(20) NOT NULL,
  `tgl` char(20) DEFAULT NULL,
  PRIMARY KEY (`nomorjurnalumum`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for totkasir
-- ----------------------------
DROP TABLE IF EXISTS `totkasir`;
CREATE TABLE `totkasir` (
  `nomorkasir` varchar(30) DEFAULT NULL,
  `subtotal` double DEFAULT NULL,
  `discount` double DEFAULT NULL,
  `total` double DEFAULT NULL,
  `tunai` double DEFAULT NULL,
  `username` varchar(100) DEFAULT NULL,
  `tgl` date DEFAULT NULL,
  `datetime` datetime DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for totmemberorder
-- ----------------------------
DROP TABLE IF EXISTS `totmemberorder`;
CREATE TABLE `totmemberorder` (
  `nomormemberorder` char(20) NOT NULL,
  `nomorpenjualan` varchar(20) DEFAULT NULL,
  `kodeanggota` char(20) DEFAULT NULL,
  `username` char(20) DEFAULT NULL,
  `kodeakun` char(20) DEFAULT NULL,
  `akunkas` varchar(20) DEFAULT NULL,
  `kodecostcenter` char(20) DEFAULT NULL,
  `fakturasli` text,
  `tgl` date DEFAULT NULL,
  `jthtmp` date DEFAULT NULL,
  `ppn` double DEFAULT NULL,
  `persdisc` double DEFAULT NULL,
  `persdisc2` double DEFAULT NULL,
  `subtotal` double DEFAULT NULL,
  `pajak` double DEFAULT NULL,
  `discount` double DEFAULT NULL,
  `discount2` double DEFAULT NULL,
  `total` double DEFAULT NULL,
  `tunai` double DEFAULT NULL,
  `piutang` double DEFAULT NULL,
  `datetime` datetime DEFAULT NULL,
  `flaglunas` varchar(1) DEFAULT NULL,
  `kodesalesman` varchar(20) DEFAULT NULL,
  `komisi` double DEFAULT NULL,
  `status` varchar(1) DEFAULT '0',
  `dp` double DEFAULT NULL,
  `jenisorder` varchar(1) DEFAULT '0',
  PRIMARY KEY (`nomormemberorder`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for totmembertopup
-- ----------------------------
DROP TABLE IF EXISTS `totmembertopup`;
CREATE TABLE `totmembertopup` (
  `nomormembertopup` char(20) NOT NULL,
  `username` char(20) NOT NULL,
  `tgl` char(20) NOT NULL,
  `status` char(20) NOT NULL DEFAULT 'D',
  PRIMARY KEY (`nomormembertopup`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for totmutasistock
-- ----------------------------
DROP TABLE IF EXISTS `totmutasistock`;
CREATE TABLE `totmutasistock` (
  `nomormutasistock` char(20) NOT NULL,
  `username` char(20) NOT NULL,
  `tgl` date DEFAULT NULL,
  `keterangan` char(200) DEFAULT NULL,
  `datetime` datetime DEFAULT NULL,
  PRIMARY KEY (`nomormutasistock`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for totpacking
-- ----------------------------
DROP TABLE IF EXISTS `totpacking`;
CREATE TABLE `totpacking` (
  `nopacking` varchar(20) NOT NULL DEFAULT '',
  `tgl` date DEFAULT NULL,
  `username` varchar(20) DEFAULT NULL,
  `keterangan` varchar(100) DEFAULT NULL,
  PRIMARY KEY (`nopacking`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for totpelunasanhutang
-- ----------------------------
DROP TABLE IF EXISTS `totpelunasanhutang`;
CREATE TABLE `totpelunasanhutang` (
  `nomorpelunasanhutang` char(20) NOT NULL,
  `username` char(20) NOT NULL,
  `kodesupplier` char(6) NOT NULL,
  `kodeakun` char(20) NOT NULL,
  `kodecostcenter` char(20) NOT NULL,
  `tgl` date DEFAULT NULL,
  `discount` double DEFAULT NULL,
  `total` double DEFAULT NULL,
  `pelunasan` double DEFAULT NULL,
  `datetime` datetime DEFAULT NULL,
  PRIMARY KEY (`nomorpelunasanhutang`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for totpelunasanpiutang
-- ----------------------------
DROP TABLE IF EXISTS `totpelunasanpiutang`;
CREATE TABLE `totpelunasanpiutang` (
  `nomorpelunasanpiutang` char(20) NOT NULL,
  `kodeanggota` char(20) NOT NULL,
  `username` char(20) NOT NULL,
  `kodeakun` char(20) NOT NULL,
  `kodecostcenter` char(20) NOT NULL,
  `tgl` date DEFAULT NULL,
  `discount` double DEFAULT NULL,
  `total` double DEFAULT NULL,
  `pelunasan` double DEFAULT NULL,
  `datetime` datetime DEFAULT NULL,
  PRIMARY KEY (`nomorpelunasanpiutang`),
  KEY `idxTotpelunasanpiutang` (`nomorpelunasanpiutang`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for totpembelian
-- ----------------------------
DROP TABLE IF EXISTS `totpembelian`;
CREATE TABLE `totpembelian` (
  `nomorpembelian` char(20) NOT NULL,
  `kodecostcenter` char(20) NOT NULL,
  `kodeakun` char(20) NOT NULL,
  `kodesupplier` char(6) NOT NULL,
  `username` char(20) NOT NULL,
  `fakturasli` char(20) DEFAULT NULL,
  `tgl` date DEFAULT NULL,
  `jthtmp` date DEFAULT NULL,
  `ppn` double DEFAULT NULL,
  `persdisc` double DEFAULT NULL,
  `persdisc2` double DEFAULT NULL,
  `subtotal` double DEFAULT NULL,
  `pajak` double DEFAULT NULL,
  `discount` double DEFAULT NULL,
  `discount2` double DEFAULT NULL,
  `total` double DEFAULT NULL,
  `tunai` double DEFAULT NULL,
  `hutang` double DEFAULT NULL,
  `datetime` datetime DEFAULT NULL,
  `flaglunas` varchar(1) DEFAULT NULL,
  `kodesalesman` varchar(20) DEFAULT NULL,
  `statuspembelian` char(1) DEFAULT '0',
  `kodegudang` varchar(20) DEFAULT NULL,
  PRIMARY KEY (`nomorpembelian`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for totpenjualan
-- ----------------------------
DROP TABLE IF EXISTS `totpenjualan`;
CREATE TABLE `totpenjualan` (
  `nomorpenjualan` char(20) NOT NULL,
  `kodeanggota` char(20) NOT NULL,
  `username` char(20) NOT NULL,
  `kodeakun` char(20) NOT NULL,
  `kodecostcenter` char(20) NOT NULL,
  `fakturasli` char(20) DEFAULT NULL,
  `tgl` date DEFAULT NULL,
  `jthtmp` date DEFAULT NULL,
  `ppn` double DEFAULT NULL,
  `persdisc` double DEFAULT NULL,
  `persdisc2` double DEFAULT NULL,
  `subtotal` double DEFAULT NULL,
  `pajak` double DEFAULT NULL,
  `discount` double DEFAULT NULL,
  `discount2` double DEFAULT NULL,
  `total` double DEFAULT NULL,
  `tunai` double DEFAULT NULL,
  `dp` double DEFAULT NULL,
  `piutang` double DEFAULT NULL,
  `datetime` datetime DEFAULT NULL,
  `flaglunas` varchar(1) DEFAULT NULL,
  `kodesalesman` varchar(20) DEFAULT NULL,
  `komisi` double DEFAULT NULL,
  `kodegudang` varchar(20) DEFAULT NULL,
  `upkepada` varchar(20) DEFAULT NULL,
  `jenis` varchar(1) DEFAULT 'R',
  `keterangan` varchar(255) NOT NULL DEFAULT '',
  PRIMARY KEY (`nomorpenjualan`),
  KEY `idxTotPenjualan` (`kodeanggota`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for totpoin
-- ----------------------------
DROP TABLE IF EXISTS `totpoin`;
CREATE TABLE `totpoin` (
  `nomortukarpoin` varchar(255) NOT NULL,
  `tgl` date DEFAULT NULL,
  `kodeanggota` varchar(255) DEFAULT NULL,
  `kodehadiah` varchar(255) DEFAULT NULL,
  `qty` double DEFAULT NULL,
  `poin` double DEFAULT NULL,
  `keterangan` varchar(255) DEFAULT NULL,
  `status` varchar(1) DEFAULT '0',
  PRIMARY KEY (`nomortukarpoin`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for totrtnpembelian
-- ----------------------------
DROP TABLE IF EXISTS `totrtnpembelian`;
CREATE TABLE `totrtnpembelian` (
  `nomorreturpembelian` char(20) NOT NULL,
  `username` char(20) NOT NULL,
  `kodesupplier` char(6) NOT NULL,
  `fakturasli` char(20) DEFAULT NULL,
  `tgl` date DEFAULT NULL,
  `jthtmp` date DEFAULT NULL,
  `ppn` double DEFAULT NULL,
  `persdisc` double DEFAULT NULL,
  `persdisc2` double DEFAULT NULL,
  `subtotal` double DEFAULT NULL,
  `pajak` double DEFAULT NULL,
  `discount` double DEFAULT NULL,
  `discount2` double DEFAULT NULL,
  `total` double DEFAULT NULL,
  `tunai` double DEFAULT NULL,
  `hutang` double DEFAULT NULL,
  `datetime` datetime DEFAULT NULL,
  `kodecostcenter` char(20) NOT NULL,
  PRIMARY KEY (`nomorreturpembelian`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for totrtnpenjualan
-- ----------------------------
DROP TABLE IF EXISTS `totrtnpenjualan`;
CREATE TABLE `totrtnpenjualan` (
  `nomorreturpenjualan` char(20) NOT NULL,
  `kodeanggota` char(20) NOT NULL,
  `username` char(20) NOT NULL,
  `fakturasli` char(20) DEFAULT NULL,
  `tgl` date DEFAULT NULL,
  `jthtmp` date DEFAULT NULL,
  `ppn` double DEFAULT NULL,
  `persdisc` double DEFAULT NULL,
  `persdisc2` double DEFAULT NULL,
  `subtotal` double DEFAULT NULL,
  `pajak` double DEFAULT NULL,
  `discount` double DEFAULT NULL,
  `discount2` double DEFAULT NULL,
  `total` double DEFAULT NULL,
  `tunai` double DEFAULT NULL,
  `piutang` double DEFAULT NULL,
  `datetime` datetime DEFAULT NULL,
  `kodecostcenter` char(20) NOT NULL,
  `nomorpenjualan` varchar(20) DEFAULT NULL,
  PRIMARY KEY (`nomorreturpenjualan`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for totstockopname
-- ----------------------------
DROP TABLE IF EXISTS `totstockopname`;
CREATE TABLE `totstockopname` (
  `nomorstockopname` char(20) NOT NULL,
  `username` char(20) NOT NULL,
  `kodecostcenter` char(20) NOT NULL DEFAULT '1',
  `kodegudang` char(20) NOT NULL,
  `tgl` date DEFAULT NULL,
  `keterangan` varchar(255) DEFAULT NULL,
  `datetime` datetime DEFAULT NULL,
  PRIMARY KEY (`nomorstockopname`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for totupdatekartupiutang
-- ----------------------------
DROP TABLE IF EXISTS `totupdatekartupiutang`;
CREATE TABLE `totupdatekartupiutang` (
  `nomorupdatekartupiutang` char(20) NOT NULL,
  `username` char(20) NOT NULL,
  `tgl` date DEFAULT NULL,
  `datetime` datetime DEFAULT NULL,
  PRIMARY KEY (`nomorupdatekartupiutang`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for tukarpoin
-- ----------------------------
DROP TABLE IF EXISTS `tukarpoin`;
CREATE TABLE `tukarpoin` (
  `nomortukarpoin` varchar(255) DEFAULT NULL,
  `faktur` varchar(255) DEFAULT NULL,
  `poin` double DEFAULT NULL,
  `tgl` date DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for updatekartupiutang
-- ----------------------------
DROP TABLE IF EXISTS `updatekartupiutang`;
CREATE TABLE `updatekartupiutang` (
  `nomorupdatekartupiutang` char(20) NOT NULL,
  `kodeanggota` varchar(40) DEFAULT NULL,
  `jumlah` double DEFAULT NULL,
  `tgl` date DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;

-- ----------------------------
-- Table structure for username
-- ----------------------------
DROP TABLE IF EXISTS `username`;
CREATE TABLE `username` (
  `kode` char(10) DEFAULT NULL,
  `username` char(20) NOT NULL,
  `userpassword` char(100) DEFAULT NULL,
  `fullname` char(100) DEFAULT NULL,
  `id` char(6) DEFAULT NULL,
  `login` char(1) DEFAULT NULL,
  `menulevel` smallint(5) unsigned DEFAULT NULL,
  PRIMARY KEY (`username`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 CHECKSUM=1 DELAY_KEY_WRITE=1 ROW_FORMAT=DYNAMIC;
DROP TRIGGER IF EXISTS `barang_masuk_insert`;
DELIMITER ;;
CREATE TRIGGER `barang_masuk_insert` AFTER INSERT ON `kartustock` FOR EACH ROW BEGIN
 UPDATE stock
 SET stok = stok + NEW.debet - NEW.kredit
 WHERE
 kodestock = NEW.kodestock;
end
;;
DELIMITER ;
DROP TRIGGER IF EXISTS `saldostock_delete`;
DELIMITER ;;
CREATE TRIGGER `saldostock_delete` AFTER DELETE ON `kartustock` FOR EACH ROW BEGIN
 UPDATE stock
 SET stok = stok + OLD.kredit - OLD.debet
 WHERE
 kodestock = OLD.kodestock;
END
;;
DELIMITER ;
