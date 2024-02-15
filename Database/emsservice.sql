ALTER TABLE `postdb`.`emsservice` 
CHANGE COLUMN `minweight` `minweight` INT NOT NULL ,
CHANGE COLUMN `maxweight` `maxweight` INT NOT NULL ,
CHANGE COLUMN `emserviceprice` `emsprice` INT NOT NULL DEFAULT 0 ,
ADD PRIMARY KEY (`minweight`, `maxweight`);
;
