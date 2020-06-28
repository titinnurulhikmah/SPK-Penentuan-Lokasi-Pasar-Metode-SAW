-- phpMyAdmin SQL Dump
-- version 4.2.11
-- http://www.phpmyadmin.net
--
-- Host: 127.0.0.1
-- Generation Time: 28 Jun 2020 pada 11.53
-- Versi Server: 5.6.21
-- PHP Version: 5.6.3

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;

--
-- Database: `lahanpasar`
--

-- --------------------------------------------------------

--
-- Struktur dari tabel `tbl_bobotkriteria`
--

CREATE TABLE IF NOT EXISTS `tbl_bobotkriteria` (
  `kd_bobot` char(10) COLLATE latin1_general_ci NOT NULL,
  `c1` int(3) NOT NULL,
  `c2` int(3) NOT NULL,
  `c3` int(3) NOT NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COLLATE=latin1_general_ci;

--
-- Dumping data untuk tabel `tbl_bobotkriteria`
--

INSERT INTO `tbl_bobotkriteria` (`kd_bobot`, `c1`, `c2`, `c3`) VALUES
('1', 25, 25, 50),
('2', 25, 25, 200);

-- --------------------------------------------------------

--
-- Struktur dari tabel `tbl_lokasi`
--

CREATE TABLE IF NOT EXISTS `tbl_lokasi` (
  `kd_lokasi` char(10) COLLATE latin1_general_ci NOT NULL,
  `nm_pasar` varchar(100) COLLATE latin1_general_ci NOT NULL,
  `alamat` varchar(50) COLLATE latin1_general_ci NOT NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COLLATE=latin1_general_ci;

--
-- Dumping data untuk tabel `tbl_lokasi`
--

INSERT INTO `tbl_lokasi` (`kd_lokasi`, `nm_pasar`, `alamat`) VALUES
('L01', 'Pulau Payung', 'jl. Pulau Payung no. 56'),
('001', 'cabbeng', 'cabbeng'),
('12', 'sds', 'vvvv');

-- --------------------------------------------------------

--
-- Struktur dari tabel `tbl_pedagang`
--

CREATE TABLE IF NOT EXISTS `tbl_pedagang` (
  `id_pedagang` char(10) COLLATE latin1_general_ci NOT NULL,
  `nm_pedagang` varchar(100) COLLATE latin1_general_ci NOT NULL,
  `tgl_registrasi` varchar(10) COLLATE latin1_general_ci NOT NULL,
  `jns_dagangan` varchar(50) COLLATE latin1_general_ci NOT NULL,
  `ukuran_kios` varchar(50) COLLATE latin1_general_ci NOT NULL,
  `no_hp` char(13) COLLATE latin1_general_ci NOT NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COLLATE=latin1_general_ci;

--
-- Dumping data untuk tabel `tbl_pedagang`
--

INSERT INTO `tbl_pedagang` (`id_pedagang`, `nm_pedagang`, `tgl_registrasi`, `jns_dagangan`, `ukuran_kios`, `no_hp`) VALUES
('p0001', 'Fitri Pratiwi', '2020-01-02', 'buah-buahan', '100', '082291937323'),
('p0002', 'Rian', '2020-03-04', 'buah', '10*10', '0822999999991');

-- --------------------------------------------------------

--
-- Struktur dari tabel `tbl_penilaian`
--

CREATE TABLE IF NOT EXISTS `tbl_penilaian` (
  `no_permohonan` char(11) COLLATE latin1_general_ci NOT NULL,
  `tgl_permohonan` date NOT NULL,
  `status` varchar(20) COLLATE latin1_general_ci NOT NULL,
  `kd_pedagang` char(10) COLLATE latin1_general_ci NOT NULL,
  `kd_lokasi` char(10) COLLATE latin1_general_ci NOT NULL,
  `kd_bobot` char(10) COLLATE latin1_general_ci NOT NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COLLATE=latin1_general_ci;

--
-- Indexes for dumped tables
--

--
-- Indexes for table `tbl_bobotkriteria`
--
ALTER TABLE `tbl_bobotkriteria`
 ADD PRIMARY KEY (`kd_bobot`);

--
-- Indexes for table `tbl_lokasi`
--
ALTER TABLE `tbl_lokasi`
 ADD PRIMARY KEY (`kd_lokasi`);

--
-- Indexes for table `tbl_pedagang`
--
ALTER TABLE `tbl_pedagang`
 ADD PRIMARY KEY (`id_pedagang`);

--
-- Indexes for table `tbl_penilaian`
--
ALTER TABLE `tbl_penilaian`
 ADD PRIMARY KEY (`no_permohonan`);

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
