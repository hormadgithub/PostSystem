ALTER TABLE `postdb`.`regrate` 
CHANGE COLUMN `minweight` `minweight` INT NOT NULL ,
CHANGE COLUMN `maxweight` `maxweight` INT NOT NULL ,
CHANGE COLUMN `regprice` `regprice` INT NOT NULL DEFAULT 0 ,
ADD PRIMARY KEY (`minweight`, `maxweight`);
;
