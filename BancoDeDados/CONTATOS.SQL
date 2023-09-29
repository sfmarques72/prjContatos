CREATE DATABASE contatos;

CREATE TABLE contatos (
	codigo integer PRIMARY KEY, 
    nome            varchar(50),
	sobrenome varchar(100),
	email varchar(255),
	telefone varchar(15)
);

insert into contatos values (1, 'JOAO', 'DA SILVA', 'joao.silva@gmail.com', '(14) 3222-2121');
insert into contatos values (2, 'MARIA', 'DOS SANTOS', 'maria.santos@gmail.com', '(14) 3222-2123');
