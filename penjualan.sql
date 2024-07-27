-- phpMyAdmin SQL Dump
-- version 3.2.4
-- http://www.phpmyadmin.net
--
-- Host: localhost
-- Generation Time: Jun 10, 2014 at 03:05 PM
-- Server version: 5.1.41
-- PHP Version: 5.3.1

SET SQL_MODE="NO_AUTO_VALUE_ON_ZERO";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;

--
-- Database: `penjualan`
--

-- --------------------------------------------------------

--
-- Table structure for table `barang`
--

CREATE TABLE IF NOT EXISTS `barang` (
  `kdbrg` varchar(5) NOT NULL,
  `nmbrg` varchar(30) NOT NULL,
  `hrgbrg` decimal(10,0) NOT NULL,
  `satuan` varchar(11) NOT NULL,
  `stok` int(11) NOT NULL,
  PRIMARY KEY (`kdbrg`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `barang`
--

INSERT INTO `barang` (`kdbrg`, `nmbrg`, `hrgbrg`, `satuan`, `stok`) VALUES
('P0001', 'Pensil', '2000', 'Buah', 1),
('P0002', 'Pulpen', '2500', 'Buah', 3),
('S0001', 'Sendok superman', '15000', 'Lusin', 8),
('S0002', 'Sendok Super', '20000', 'Lusin', 3),
('T0001', 'Tempat sampah', '15000', 'Buah', 0),
('S0003', 'Sendok bukan sembarang sendok', '30000', 'Lusin', 10),
('P0003', 'Pensil murah meriah', '1500', 'Lusin', 0),
('P0004', 'Pulpen sayang anak', '1250', 'Buah', 4),
('R0001', 'Rokok tjap Aki-Aki', '12000', 'Bungkus', 10);

-- --------------------------------------------------------

--
-- Table structure for table `customer`
--

CREATE TABLE IF NOT EXISTS `customer` (
  `kdcust` varchar(10) NOT NULL,
  `nmcust` varchar(30) NOT NULL,
  `alamat` varchar(40) NOT NULL,
  `jenkel` varchar(10) NOT NULL,
  `nohp` varchar(15) NOT NULL,
  `tglmsk` varchar(10) NOT NULL,
  PRIMARY KEY (`kdcust`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `customer`
--

INSERT INTO `customer` (`kdcust`, `nmcust`, `alamat`, `jenkel`, `nohp`, `tglmsk`) VALUES
('A0001', 'Andi Malalari', 'Jl. Keder 13', 'Laki-laki', '081588378553', '12/5/2013'),
('A0002', 'Anton Semarank', 'Jl. semanggi 13 jawa', 'Laki-laki', '083847575757', '12/2/2013'),
('A0003', 'Panjul bin peang', 'Jl. mengor 13', 'Laki-laki', '088888888', '12/4/2013');

-- --------------------------------------------------------

--
-- Table structure for table `detailtrx`
--

CREATE TABLE IF NOT EXISTS `detailtrx` (
  `kdtrx` varchar(10) NOT NULL,
  `kdbrg` varchar(5) NOT NULL,
  `qty` int(11) NOT NULL,
  `disc` smallint(6) NOT NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `detailtrx`
--

INSERT INTO `detailtrx` (`kdtrx`, `kdbrg`, `qty`, `disc`) VALUES
('TX0001', 'p0003', 200, 10),
('TX0001', 's0001', 100, 16),
('TX0001', 's0002', 20, 0),
('TX0002', 'p0001', 10, 0),
('TX0002', 's0002', 20, 1),
('TX0002', 'p0003', 10, 0),
('TX0002', 's0001', 5, 0),
('TX0002', 's0003', 100, 2),
('TX0003', 'p0003', 20, 1),
('TX0003', 's0003', 100, 10),
('TX0003', 'p0001', 10, 0),
('TX0004', 's0002', 5, 0),
('TX0004', 'p0002', 2, 0),
('TX0004', 't0001', 3, 0),
('TX0005', 's0002', 6, 2),
('TX0005', 'p0003', 5, 1),
('TX0005', 't0001', 1, 0),
('TX0005', 'p0001', 5, 3),
('TX0006', 'p0001', 2, 0),
('TX0006', 's0002', 5, 2),
('TX0006', 's0003', 20, 10),
('TX0007', 'p0002', 2, 0),
('TX0007', 's0002', 5, 1),
('TX0008', 's0003', 20, 5),
('TX0008', 's0002', 2, 0),
('TX0008', 't0001', 1, 0);

-- --------------------------------------------------------

--
-- Table structure for table `transaksi`
--

CREATE TABLE IF NOT EXISTS `transaksi` (
  `kdtrx` varchar(10) NOT NULL,
  `tgltrx` date NOT NULL,
  `kdcust` varchar(10) NOT NULL,
  PRIMARY KEY (`kdtrx`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `transaksi`
--

INSERT INTO `transaksi` (`kdtrx`, `tgltrx`, `kdcust`) VALUES
('TX0002', '2013-12-10', 'A0001'),
('TX0001', '2013-12-08', 'A0002'),
('TX0003', '2013-12-10', 'A0001'),
('TX0004', '2013-12-10', 'A0002'),
('TX0005', '2013-12-14', 'A0001'),
('TX0006', '2013-12-14', 'A0002'),
('TX0007', '2013-12-17', 'A0003'),
('TX0008', '2013-12-17', 'A0002');

-- --------------------------------------------------------

--
-- Table structure for table `user`
--

CREATE TABLE IF NOT EXISTS `user` (
  `userid` varchar(10) NOT NULL,
  `nmuser` varchar(20) NOT NULL,
  `pass` varchar(20) NOT NULL,
  PRIMARY KEY (`userid`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `user`
--

INSERT INTO `user` (`userid`, `nmuser`, `pass`) VALUES
('admin', 'admin', 'admin');

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
