using System;

class Colaborador : Pessoa
{
    // Atributo
    private double salario;
    // COnstrutor
    public Colaborador(string nome, int idade, double salario)
    {
        this.nome = nome;
        this.idade = idade;
        this.salario = salario;
        mensagemPessoa();
        mensagemColaborador();
    }
    // Metodos
    private void mensagemColaborador()
    {
        Console.WriteLine("Salario: " + salario);  
    }
}