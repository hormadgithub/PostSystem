USE POSTDB;


CREATE TABLE `customer` (
  `CSCode` varchar(10) CHARACTER SET utf8 COLLATE utf8_unicode_ci NOT NULL DEFAULT '',
  `CSName` varchar(100) CHARACTER SET utf8 COLLATE utf8_unicode_ci NOT NULL DEFAULT '',
  `Address` varchar(200) CHARACTER SET utf8 COLLATE utf8_unicode_ci NOT NULL DEFAULT '',
  `Tel` varchar(50) CHARACTER SET utf8 COLLATE utf8_unicode_ci NOT NULL DEFAULT '',
  `LastUpdate` datetime DEFAULT NULL,
  `LastUser` varchar(10) COLLATE utf8_unicode_ci DEFAULT NULL,
  PRIMARY KEY (`CSCode`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;

CREATE TABLE `emsrate` (
  `minweight` int(11) NOT NULL DEFAULT 0,
  `maxweight` int(11) NOT NULL DEFAULT 0,
  `emsprice` int(11) NOT NULL DEFAULT 0,
  `emsservice` int(11) DEFAULT 0,
  `LastUpdate` datetime DEFAULT NULL,
  `LastUser` varchar(10) COLLATE utf8_unicode_ci DEFAULT NULL,
  PRIMARY KEY (`minweight`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;


CREATE TABLE `regrate` (
  `minweight` int(11) NOT NULL DEFAULT 0,
  `maxweight` int(11) NOT NULL DEFAULT 0,
  `regprice` int(11) DEFAULT 0,
  `regservice` int(11) DEFAULT 0,
  `LastUpdate` datetime DEFAULT NULL,
  `LastUser` varchar(10) COLLATE utf8_unicode_ci DEFAULT NULL,
  PRIMARY KEY (`minweight`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;


CREATE TABLE `users` (
  `userid` varchar(10) COLLATE utf8_unicode_ci NOT NULL,
  `password` varchar(15) COLLATE utf8_unicode_ci DEFAULT NULL,
  `LastUpdate` datetime DEFAULT NULL,
  `LastUser` varchar(10) COLLATE utf8_unicode_ci DEFAULT NULL,
  PRIMARY KEY (`userid`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;


CREATE TABLE `invoice` (
  `Invno` varchar(10) COLLATE utf8_unicode_ci NOT NULL,
  `Invdate` date DEFAULT NULL,
  `CSCode` varchar(10) COLLATE utf8_unicode_ci NOT NULL,
  `LastUpdate` datetime DEFAULT NULL,
  `LastUser` varchar(10) COLLATE utf8_unicode_ci DEFAULT NULL,
  PRIMARY KEY (`Invno`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;

CREATE TABLE `invoicedetail` (
  `invno` varchar(10) COLLATE utf8_unicode_ci NOT NULL,
  `invdt_item` int(11) NOT NULL,
  `sendtype` varchar(3) COLLATE utf8_unicode_ci NOT NULL,
  `weight` int(11) NOT NULL DEFAULT 0,
  `unitprice` decimal(10,0) DEFAULT 0,
  `serviceprice` decimal(10,0) DEFAULT 0,
  `trackno` varchar(10) COLLATE utf8_unicode_ci DEFAULT NULL,
  `LastUpdate` datetime DEFAULT NULL,
  `LastUser` varchar(10) COLLATE utf8_unicode_ci DEFAULT NULL,
  PRIMARY KEY (`invno`,`invdt_item`),
  CONSTRAINT `fk_invno` FOREIGN KEY (`invno`) REFERENCES `invoice` (`Invno`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;
