<?php

namespace Tesla\InsertQuestionsToolaft\lib;

use PDO;
use PDOException;

class Database{
    private string $host;
    private string $db;
    private string $user;
    private string $password;
    private string $charset;
    private string $port;

    public function __construct(){
        $this->host = $_ENV['HOST'];
        $this->db = $_ENV['DB'];
        $this->user = $_ENV['USER'];
        $this->password = $_ENV['PASSWORD'];
        $this->charset = $_ENV['CHATSET'];
        $this->port = $_ENV['PORT'];
    }

    public function connect(){
        try {

            $connection = "mysql:host=" . $this->host . ";port=" . $this->port . ";dbname=" . $this->db . ";charset=" . $this->charset;

            $options = [
                PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
                PDO::ATTR_EMULATE_PREPARES => false
            ];

            $pdo = new PDO(
                $connection, 
                $this->user, 
                $this->password, 
                $options
            );
            
            return $pdo;

        } catch (PDOException $error) {
            
            print_r("Error connection: ". $error->getMessage());

        }
    }
}