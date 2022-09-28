using System;
namespace _03Conceitos {
class Aluno
{
    // Atributos
    public string nome;
    public double nota1, nota2;
    // Media
    public double media() {
        return (nota1 + nota2) / 2;
    }
    // Situação
    public string situacao(double media) {
        return media >= 7 ? "aprovado" : "reprovado";
    }
    //mensagem
    public void mensagem() {
        double obterMedia = media();
        string obterSituacao = situacao(obterMedia);
        Console.WriteLine(nome+ " esta "+ obterSituacao+ " com media: " + obterMedia);
    }
}
}