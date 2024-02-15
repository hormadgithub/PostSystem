ALTER TABLE `postdb`.`regservice` 
CHANGE COLUMN `minweight` `minweight` INT NOT NULL ,
CHANGE COLUMN `maxweight` `maxweight` INT NOT NULL ,
CHANGE COLUMN `regerviceprice` `regprice` INT NOT NULL DEFAULT 0 ,
ADD PRIMARY KEY (`minweight`, `maxweight`);
;
