-- phpMyAdmin SQL Dump
-- version 5.1.1
-- https://www.phpmyadmin.net/
--
-- Host: 127.0.0.1
-- Generation Time: Feb 08, 2024 at 03:47 AM
-- Server version: 10.4.22-MariaDB
-- PHP Version: 8.1.2

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Database: `futureskill`
--

-- --------------------------------------------------------

--
-- Table structure for table `customer`
--

CREATE TABLE `customer` (
  `Customer_ID` varchar(8) COLLATE utf8mb4_thai_520_w2 NOT NULL,
  `Fname` varchar(20) COLLATE utf8mb4_thai_520_w2 NOT NULL,
  `Lname` varchar(20) COLLATE utf8mb4_thai_520_w2 NOT NULL,
  `Address` varchar(40) COLLATE utf8mb4_thai_520_w2 DEFAULT NULL,
  `Post_code` varchar(5) COLLATE utf8mb4_thai_520_w2 NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_thai_520_w2;

--
-- Dumping data for table `customer`
--

INSERT INTO `customer` (`Customer_ID`, `Fname`, `Lname`, `Address`, `Post_code`) VALUES
('1', 'Flor', 'Stone', NULL, '10520'),
('10', 'Charlie', 'Sutherland', NULL, '31000'),
('100', 'Elisha', 'Lloyd', NULL, '10540'),
('101', 'Samara', 'Barnett', NULL, '50000'),
('102', 'Nadia', 'Sawyer', NULL, '10540'),
('103', 'Melita', 'Holcomb', NULL, '52000'),
('104', 'Evita', 'Dickerson', NULL, '57000'),
('105', 'Aundrea', 'Barry', NULL, '52000'),
('106', 'Irvin', 'Leach', NULL, '52000'),
('107', 'Kellee', 'Molina', NULL, '20000'),
('108', 'Ronny', 'Sykes', NULL, '50000'),
('109', 'Ocie', 'Walton', NULL, '10520'),
('11', 'Charlie', 'Pacino', NULL, '52000'),
('110', 'Reva', 'Fuller', NULL, '57000'),
('111', 'Marica', 'Henderson', NULL, '10540'),
('112', 'Vanetta', 'Gross', NULL, '81000'),
('113', 'Audrea', 'Hayden', NULL, '57000'),
('114', 'Neda', 'Mcfadden', NULL, '60000'),
('115', 'Dell', 'Wilkinson', NULL, '20000'),
('116', 'Sharee', 'Carver', NULL, '60000'),
('117', 'Joaquina', 'Mcconnell', NULL, '31000'),
('118', 'Carlyn', 'Calderon', NULL, '60000'),
('119', 'Olga', 'Wilkins', NULL, '40000'),
('12', 'Guillaume', 'Jackson', NULL, '57000'),
('120', 'Delpha', 'Golden', NULL, '10540'),
('121', 'Sharell', 'Branch', NULL, '81000'),
('122', 'Eleanor', 'Farley', NULL, '52000'),
('123', 'Justina', 'Hayes', NULL, '10520'),
('124', 'Omer', 'Macdonald', NULL, '60000'),
('125', 'Roseline', 'Hill', NULL, '57000'),
('126', 'Janeen', 'Holt', NULL, '81000'),
('127', 'Tory', 'Avila', NULL, '60000'),
('128', 'Keila', 'Slater', NULL, '31000'),
('129', 'Jazmin', 'Burch', NULL, '10540'),
('13', 'Daniel', 'Costner', NULL, '50000'),
('130', 'Tandy', 'Ramsey', NULL, '20000'),
('131', 'Debra', 'Herring', NULL, '50000'),
('132', 'Monnie', 'Walker', NULL, '40000'),
('133', 'Alvaro', 'Hooper', NULL, '52000'),
('134', 'Leandro', 'West', NULL, '52000'),
('135', 'Lurline', 'Dennis', NULL, '20000'),
('136', 'Aurea', 'Sanders', NULL, '20000'),
('137', 'Stacia', 'Rutledge', NULL, '10540'),
('138', 'Philomena', 'Elliott', NULL, '20000'),
('139', 'Gonzalo', 'Stone', NULL, '10540'),
('14', 'Dianne', 'Derek', NULL, '20000'),
('140', 'So', 'Hardin', NULL, '60000'),
('141', 'Lauren', 'Rodriquez', NULL, '50000'),
('142', 'Loan', 'Patton', NULL, '52000'),
('143', 'Ha', 'Lowe', NULL, '57000'),
('144', 'Alexandra', 'Mcgowan', NULL, '20000'),
('145', 'Melissa', 'Nichols', NULL, '81000'),
('146', 'Keesha', 'Cohen', NULL, '31000'),
('147', 'Elvera', 'Benson', NULL, '31000'),
('148', 'Lanora', 'Ray', NULL, '40000'),
('149', 'Jimmy', 'Weber', NULL, '20000'),
('15', 'Geraldine', 'Schneider', NULL, '57000'),
('150', 'Mae', 'Wilkinson', NULL, '10520'),
('151', 'Lino', 'Noble', NULL, '10540'),
('152', 'Kathie', 'Jordan', NULL, '57000'),
('153', 'Heike', 'Emerson', NULL, '10520'),
('154', 'Princess', 'Jones', NULL, '60000'),
('155', 'Doretha', 'Tyler', NULL, '57000'),
('156', 'Emilio', 'Duncan', NULL, '10520'),
('157', 'Glinda', 'Lambert', NULL, '60000'),
('158', 'Emery', 'Schneider', NULL, '31000'),
('159', 'Amado', 'Jefferson', NULL, '52000'),
('16', 'Geraldine', 'Martin', NULL, '31000'),
('160', 'Hisako', 'Herrera', NULL, '52000'),
('161', 'Jamee', 'Mclaughlin', NULL, '50000'),
('162', 'Audrie', 'Cannon', NULL, '60000'),
('163', 'Lacey', 'Walters', NULL, '10520'),
('164', 'Weldon', 'Robinson', NULL, '20000'),
('165', 'Solange', 'Gates', NULL, '57000'),
('166', 'Jame', 'Calhoun', NULL, '20000'),
('167', 'Stefan', 'Walters', NULL, '60000'),
('168', 'Arlene', 'Elliott', NULL, '57000'),
('169', 'Sofia', 'Burnett', NULL, '50000'),
('17', 'Guillaume', 'Edwards', NULL, '50000'),
('170', 'Christal', 'Grant', NULL, '31000'),
('171', 'Violeta', 'Stokes', NULL, '52000'),
('172', 'Jackeline', 'Finley', NULL, '57000'),
('173', 'Heidy', 'Massey', NULL, '40000'),
('174', 'Yolando', 'Wilkerson', NULL, '52000'),
('175', 'Lucinda', 'Hurley', NULL, '57000'),
('176', 'Gita', 'Logan', NULL, '10540'),
('177', 'Felicita', 'Alston', NULL, '52000'),
('178', 'Lela', 'Ewing', NULL, '57000'),
('179', 'Gabriel', 'Pollard', NULL, '60000'),
('18', 'Maurice', 'Mahoney', NULL, '50000'),
('180', 'Vincenza', 'Walton', NULL, '10540'),
('181', 'Shandra', 'Higgins', NULL, '50000'),
('182', 'Thomasena', 'Pickett', NULL, '10540'),
('183', 'Aaron', 'Holder', NULL, '81000'),
('184', 'Carlos', 'Moody', NULL, '10520'),
('185', 'Javier', 'Kim', NULL, '81000'),
('186', 'Fran', 'Battle', NULL, '31000'),
('187', 'Shelia', 'Brewer', NULL, '57000'),
('188', 'Jeri', 'Randall', NULL, '57000'),
('189', 'Jermaine', 'Cote', NULL, '50000'),
('19', 'Maurice', 'Hasan', NULL, '10520'),
('190', 'Kenneth', 'Simmons', NULL, '10540'),
('191', 'Dulce', 'Matthews', NULL, '52000'),
('192', 'Kristle', 'Cain', NULL, '81000'),
('193', 'Erika', 'Murray', NULL, '60000'),
('194', 'Maud', 'Cohen', NULL, '57000'),
('195', 'Rolanda', 'Vang', NULL, '57000'),
('196', 'Renea', 'Carney', NULL, '81000'),
('197', 'Sybil', 'Harding', NULL, '40000'),
('198', 'Colette', 'Estrada', NULL, '57000'),
('199', 'Maximo', 'Ortega', NULL, '20000'),
('2', 'Lavera', 'Emerson', NULL, '57000'),
('20', 'Diane', 'Higgins', NULL, '52000'),
('200', 'Kina', 'Larsen', NULL, '57000'),
('201', 'Jarvis', 'Allison', NULL, '31000'),
('202', 'Collene', 'Newton', NULL, '52000'),
('203', 'Vella', 'Hancock', NULL, '10520'),
('204', 'Retta', 'Martinez', NULL, '10520'),
('205', 'Annelle', 'Lawrence', NULL, '52000'),
('206', 'Sherron', 'Simon', NULL, '20000'),
('207', 'Carita', 'Mcintyre', NULL, '52000'),
('208', 'Stephaine', 'Booker', NULL, '10520'),
('209', 'Emilie', 'Parsons', NULL, '31000'),
('21', 'Dianne', 'Sen', NULL, '31000'),
('210', 'Jaleesa', 'Bowen', NULL, '57000'),
('211', 'Jeannie', 'Poole', NULL, '81000'),
('212', 'Adrienne', 'Lang', NULL, '81000'),
('213', 'Jess', 'Nguyen', NULL, '10540'),
('214', 'Tandy', 'House', NULL, '52000'),
('215', 'Herman', 'Stokes', NULL, '31000'),
('216', 'Keesha', 'Lambert', NULL, '50000'),
('217', 'Lauren', 'Williamson', NULL, '31000'),
('218', 'Lucius', 'Abbott', NULL, '40000'),
('219', 'Beatrice', 'Ford', NULL, '10520'),
('22', 'Maurice', 'Daltrey', NULL, '10540'),
('220', 'Tamisha', 'Vargas', NULL, '31000'),
('221', 'Dick', 'Lamb', NULL, '81000'),
('222', 'Yolanda', 'Ball', NULL, '10520'),
('223', 'Jerica', 'Brooks', NULL, '10520'),
('224', 'Nichol', 'Carter', NULL, '50000'),
('225', 'Bobby', 'Wilson', NULL, '50000'),
('226', 'Eve', 'Saunders', NULL, '60000'),
('227', 'Ginger', 'Atkinson', NULL, '40000'),
('228', 'Kenton', 'Harrell', NULL, '40000'),
('229', 'Misti', 'Velazquez', NULL, '10520'),
('23', 'Tess', 'Roth', NULL, '20000'),
('230', 'Brandie', 'Buchanan', NULL, '60000'),
('231', 'Kandi', 'Holden', NULL, '20000'),
('232', 'Lashonda', 'Cunningham', NULL, '31000'),
('233', 'Roxanne', 'Michalkow', NULL, '31000'),
('234', 'Phuong', 'Sanders', NULL, '31000'),
('235', 'Vida', 'Kline', NULL, '10520'),
('236', 'Mayola', 'Houston', NULL, '52000'),
('237', 'Cristine', 'Bell', NULL, '81000'),
('238', 'Shamika', 'Bauer', NULL, '81000'),
('239', 'Rebbeca', 'Day', NULL, '31000'),
('24', 'Ka', 'Kaufman', NULL, '10540'),
('240', 'Maryrose', 'Cain', NULL, '40000'),
('241', 'Johnie', 'Rojas', NULL, '10520'),
('242', 'Verla', 'Christian', NULL, '81000'),
('243', 'Teressa', 'Hewitt', NULL, '60000'),
('244', 'Shenna', 'Blair', NULL, '40000'),
('245', 'Latoyia', 'Burt', NULL, '10540'),
('246', 'Shaquita', 'Norman', NULL, '10540'),
('247', 'Myrna', 'Romero', NULL, '31000'),
('248', 'Bill', 'Stein', NULL, '40000'),
('249', 'Olivia', 'Conrad', NULL, '20000'),
('25', 'Sharyl', 'Montoya', NULL, '60000'),
('250', 'Marianne', 'Bryant', NULL, '60000'),
('251', 'Dustin', 'Paul', NULL, '20000'),
('252', 'Elroy', 'Wolf', NULL, '10520'),
('253', 'Jadwiga', 'Morrison', NULL, '31000'),
('254', 'Isabella', 'Ortega', NULL, '31000'),
('255', 'Katia', 'Case', NULL, '50000'),
('256', 'Enriqueta', 'Odom', NULL, '60000'),
('257', 'Don', 'Hansen', NULL, '81000'),
('258', 'Herb', 'Sloan', NULL, '60000'),
('259', 'Blanche', 'Robbins', NULL, '20000'),
('26', 'Daniel', 'Glass', NULL, '20000'),
('260', 'Gaynell', 'Burton', NULL, '52000'),
('261', 'Marhta', 'Baldwin', NULL, '31000'),
('262', 'Peter', 'Gamble', NULL, '81000'),
('263', 'Nathan', 'Callahan', NULL, '50000'),
('264', 'Tereasa', 'Padilla', NULL, '50000'),
('265', 'Merrilee', 'Phillips', NULL, '10540'),
('266', 'Tisha', 'Mcdonald', NULL, '40000'),
('267', 'Napoleon', 'Cruz', NULL, '50000'),
('268', 'Basilia', 'Downs', NULL, '81000'),
('269', 'Alessandra', 'Estrada', NULL, '10520'),
('27', 'Rena', 'Arnold', NULL, '60000'),
('270', 'Sherron', 'Flores', NULL, '10540'),
('271', 'Elicia', 'Townsend', NULL, '81000'),
('272', 'Percy', 'Hernandez', NULL, '40000'),
('273', 'Shelia', 'Cain', NULL, '10540'),
('274', 'Charlene', 'Booker', NULL, '81000'),
('275', 'Mac', 'Mckay', NULL, '10520'),
('276', 'Jamison', 'Merritt', NULL, '50000'),
('277', 'Quinton', 'Woods', NULL, '31000'),
('278', 'Cathey', 'Mcdowell', NULL, '81000'),
('279', 'Deneen', 'Hays', NULL, '20000'),
('28', 'Arlyne', 'Ingram', NULL, '60000'),
('280', 'Jackson', 'Morin', NULL, '57000'),
('281', 'Margo', 'Hoffman', NULL, '50000'),
('282', 'Adah', 'Myers', NULL, '31000'),
('283', 'Florene', 'Craig', NULL, '52000'),
('284', 'Janey', 'Burris', NULL, '57000'),
('285', 'Evelina', 'Campbell', NULL, '40000'),
('286', 'Thi', 'Maxwell', NULL, '57000'),
('287', 'Aleshia', 'Reese', NULL, '31000'),
('288', 'Dorotha', 'Wong', NULL, '20000'),
('289', 'Estela', 'Donovan', NULL, '10520'),
('29', 'Willie', 'Barrera', NULL, '31000'),
('290', 'Hassan', 'Rivers', NULL, '10520'),
('291', 'Jen', 'Mcmahon', NULL, '20000'),
('292', 'Barbie', 'Carter', NULL, '81000'),
('293', 'Cleo', 'English', NULL, '81000'),
('294', 'Liana', 'Pena', NULL, '40000'),
('295', 'Francina', 'Slater', NULL, '52000'),
('296', 'Josie', 'Steele', NULL, '60000'),
('297', 'Maryrose', 'Roberson', NULL, '31000'),
('298', 'Elly', 'Salinas', NULL, '10540'),
('299', 'Ginny', 'Carlson', NULL, '52000'),
('3', 'Fern', 'Head', NULL, '57000'),
('30', 'Mireya', 'Cochran', NULL, '20000'),
('300', 'Mozelle', 'Schneider', NULL, '50000'),
('301', 'Adam', 'Jacobs', NULL, '50000'),
('302', 'Vernia', 'Hayes', NULL, '57000'),
('303', 'Ismael', 'Solomon', NULL, '81000'),
('304', 'Ivey', 'Rutledge', NULL, '50000'),
('305', 'Jame', 'Terrell', NULL, '10540'),
('306', 'Daina', 'Combs', NULL, '40000'),
('307', 'Lashunda', 'Davidson', NULL, '10540'),
('308', 'Arlette', 'Thornton', NULL, '10540'),
('309', 'Darron', 'Robertson', NULL, '10540'),
('31', 'Marlene', 'Odom', NULL, '10540'),
('310', 'Maple', 'Barnett', NULL, '10520'),
('311', 'Charlsie', 'Carey', NULL, '10540'),
('312', 'Frank', 'Shannon', NULL, '57000'),
('313', 'Josiah', 'Beasley', NULL, '50000'),
('314', 'Annabelle', 'Butler', NULL, '57000'),
('315', 'Ed', 'Mueller', NULL, '20000'),
('316', 'Colleen', 'Estrada', NULL, '50000'),
('317', 'Nikia', 'Kent', NULL, '31000'),
('318', 'Amber', 'Brady', NULL, '31000'),
('319', 'Wendell', 'Massey', NULL, '81000'),
('32', 'Jaclyn', 'Atkinson', NULL, '31000'),
('33', 'Al', 'Schultz', NULL, '60000'),
('34', 'Felicitas', 'Riley', NULL, '52000'),
('35', 'Cora', 'Calhoun', NULL, '20000'),
('36', 'Trula', 'Buckley', NULL, '31000'),
('37', 'Sasha', 'Wallace', NULL, '10520'),
('38', 'Caitlin', 'Hill', NULL, '40000'),
('39', 'Gino', 'Pickett', NULL, '52000'),
('4', 'Shyla', 'Ortiz', NULL, '31000'),
('40', 'Amira', 'Macdonald', NULL, '50000'),
('41', 'Mack', 'Morse', NULL, '81000'),
('42', 'Eboni', 'Jarvis', NULL, '31000'),
('43', 'Gabrielle', 'Dennis', NULL, '40000'),
('44', 'Classie', 'Norris', NULL, '20000'),
('45', 'Khalilah', 'Holman', NULL, '81000'),
('46', 'Isreal', 'Rose', NULL, '50000'),
('47', 'Verena', 'Hopper', NULL, '57000'),
('48', 'Audie', 'Flores', NULL, '10520'),
('49', 'Gertrude', 'Cooke', NULL, '60000'),
('5', 'Jeni', 'Levy', NULL, '60000'),
('50', 'Princess', 'Kane', NULL, '10520'),
('51', 'Jacinta', 'Faulkner', NULL, '20000'),
('52', 'Lashon', 'Wall', NULL, '10540'),
('53', 'Corliss', 'Mcneil', NULL, '50000'),
('54', 'Brock', 'Webb', NULL, '57000'),
('55', 'Melany', 'Mcmahon', NULL, '31000'),
('56', 'Rogelio', 'Kirby', NULL, '52000'),
('57', 'Wallace', 'Dillard', NULL, '10540'),
('58', 'Gia', 'Boyle', NULL, '31000'),
('59', 'Tennie', 'Gay', NULL, '81000'),
('6', 'Matthias', 'Hannah', NULL, '52000'),
('60', 'Ophelia', 'Hurst', NULL, '20000'),
('61', 'Ciara', 'Flowers', NULL, '60000'),
('62', 'Birgit', 'Stephenson', NULL, '10520'),
('63', 'Shameka', 'Spencer', NULL, '40000'),
('64', 'Alysa', 'Kane', NULL, '10540'),
('65', 'Fransisca', 'Reeves', NULL, '57000'),
('66', 'Jessika', 'Merritt', NULL, '20000'),
('67', 'Agustina', 'Conner', NULL, '10540'),
('68', 'Roxanna', 'Wade', NULL, '52000'),
('69', 'Virgie', 'Mays', NULL, '57000'),
('7', 'Matthias', 'Cruise', NULL, '10520'),
('70', 'Trang', 'Mcconnell', NULL, '50000'),
('71', 'Nada', 'West', NULL, '60000'),
('72', 'Catherina', 'Haney', NULL, '52000'),
('73', 'Harriette', 'Melton', NULL, '10520'),
('74', 'Willette', 'Rodgers', NULL, '10520'),
('75', 'Calandra', 'Williamson', NULL, '52000'),
('76', 'Tamatha', 'Delgado', NULL, '40000'),
('77', 'Felix', 'Ferguson', NULL, '52000'),
('78', 'Elwood', 'Hampton', NULL, '31000'),
('79', 'Josh', 'Roach', NULL, '20000'),
('8', 'Meenakshi', 'Mason', NULL, '52000'),
('80', 'Luanna', 'Scott', NULL, '50000'),
('81', 'Meryl', 'Cole', NULL, '81000'),
('82', 'Jannet', 'Elliott', NULL, '60000'),
('83', 'Fae', 'Glenn', NULL, '40000'),
('84', 'Francisco', 'Cummings', NULL, '40000'),
('85', 'Ermelinda', 'Benton', NULL, '20000'),
('86', 'Lasonya', 'Beard', NULL, '52000'),
('87', 'Rayna', 'Tran', NULL, '60000'),
('88', 'Annice', 'Boyer', NULL, '31000'),
('89', 'Ja', 'Whitfield', NULL, '31000'),
('9', 'Christian', 'Cage', NULL, '52000'),
('90', 'Jaime', 'Lester', NULL, '31000'),
('91', 'Charlsie', 'Lindsey', NULL, '60000'),
('92', 'Jannette', 'Henry', NULL, '20000'),
('93', 'Margart', 'Mccall', NULL, '40000'),
('94', 'Twanna', 'Cleveland', NULL, '20000'),
('95', 'Bronwyn', 'Horn', NULL, '57000'),
('96', 'Jon', 'Petersen', NULL, '60000'),
('97', 'Denny', 'Daniel', NULL, '10540'),
('98', 'Avis', 'Moore', NULL, '50000'),
('99', 'Eden', 'Burke', NULL, '10520');

--
-- Indexes for dumped tables
--

--
-- Indexes for table `customer`
--
ALTER TABLE `customer`
  ADD PRIMARY KEY (`Customer_ID`),
  ADD KEY `Post_code` (`Post_code`);

--
-- Constraints for dumped tables
--

--
-- Constraints for table `customer`
--
ALTER TABLE `customer`
  ADD CONSTRAINT `Customer_ibfk_1` FOREIGN KEY (`Post_code`) REFERENCES `post_code` (`Post_code`);
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
