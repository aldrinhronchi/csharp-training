using System;
class Pessoa
{
    // atributo
    private string nome = "Teste";
    // contrutor
    public Pessoa(string nome)
    {
        Console.WriteLine(nome);
        Console.WriteLine(this.nome);
    }
}