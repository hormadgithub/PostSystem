CREATE TABLE `emsrate` (
  `minweight` int NOT NULL,
  `maxweight` int NOT NULL,
  `emsprice` int NOT NULL DEFAULT '0',
  PRIMARY KEY (`minweight`,`maxweight`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;