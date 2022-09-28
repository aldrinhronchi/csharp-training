using System;

class Pessoa
{
            // metodo 1
    public void apresentar() {
        Console.WriteLine("Ola!!");
    }
            // metodo 2
    public void apresentar(string nome) {
        Console.WriteLine("Ola "+nome );
    }
            // metodo 3
    public void apresentar(string nome, int idade) {
        Console.WriteLine("Ola "+nome+" voce tem "+idade+" anos");  
    }
}