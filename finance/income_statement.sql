-- phpMyAdmin SQL Dump
-- version 4.1.8
-- http://www.phpmyadmin.net
--
-- Host: localhost
-- Generation Time: Jul 21, 2014 at 10:50 AM
-- Server version: 5.1.68-cll-lve
-- PHP Version: 5.4.23

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;

--
-- Database: `softwa19_wp59`
--

-- --------------------------------------------------------

--
-- Table structure for table `income_statement`
--

CREATE TABLE IF NOT EXISTS `income_statement` (
  `finyear` int(11) DEFAULT NULL,
  `company` varchar(100) DEFAULT NULL,
  `totalrevenue_sales` int(11) DEFAULT NULL,
  `costrevenue` int(11) DEFAULT NULL,
  `grossprofit` int(11) DEFAULT NULL,
  `generaladmexpenses` int(11) DEFAULT NULL,
  `researchdevelop` int(11) DEFAULT NULL,
  `salesmarketingexp` int(11) DEFAULT NULL,
  `depreciationamort` int(11) DEFAULT NULL,
  `interestexp1` int(11) DEFAULT NULL,
  `businespecific1` int(11) DEFAULT NULL,
  `businespecific2` int(11) DEFAULT NULL,
  `businespecific3` int(11) DEFAULT NULL,
  `miscexp` int(11) DEFAULT NULL,
  `operatingexp` int(11) DEFAULT NULL,
  `operatinginc` int(11) DEFAULT NULL,
  `interestexp2` int(11) DEFAULT NULL,
  `incomebeforetax` int(11) DEFAULT NULL,
  `tax` int(11) DEFAULT NULL,
  `netprofit` int(11) DEFAULT NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `income_statement`
--

INSERT INTO `income_statement` (`finyear`, `company`, `totalrevenue_sales`, `costrevenue`, `grossprofit`, `generaladmexpenses`, `researchdevelop`, `salesmarketingexp`, `depreciationamort`, `interestexp1`, `businespecific1`, `businespecific2`, `businespecific3`, `miscexp`, `operatingexp`, `operatinginc`, `interestexp2`, `incomebeforetax`, `tax`, `netprofit`) VALUES
(2003, 'ABC Ltd.', 123000, 93396, 29604, 7053, 0, 250, 0, 908, 0, 0, 0, 2000, 10211, 19393, 209, 19184, 3954, 15230),
(2004, 'ABC Ltd.', 147000, 81916, 65084, 7215, 0, 123, 0, 986, 0, 0, 0, 1500, 9824, 55260, 122, 55138, 4035, 51103),
(2005, 'ABC Ltd.', 125000, 73375, 51625, 6974, 0, 125, 0, 728, 0, 0, 0, 3200, 11027, 40598, 154, 40444, 3696, 36748),
(2006, 'ABC Ltd.', 165000, 61665, 103335, 6821, 0, 550, 0, 308, 0, 0, 0, 1540, 9219, 94116, 200, 93916, 4056, 89860),
(2007, 'ABC Ltd.', 190000, 63007, 126993, 20692, 0, 325, 0, 326, 0, 0, 0, 2310, 23653, 103340, 206, 103134, 3790, 99344);

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
