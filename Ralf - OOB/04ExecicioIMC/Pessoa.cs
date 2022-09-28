using System;
namespace _04Conceitos {
class Pessoa
{
    // Atributos
    public string nome;
    public double peso;
    public double altura;

    public double calculoIMC() {
        double imc = peso / (altura * altura);
        return imc;
    }
    public string retorno() {
        string ret;
        double imc = calculoIMC();
        if (imc < 18.5) {
            ret = "< 18.5 - Abaixo do peso";
        } else if (imc < 25 ) {
            ret = "< 25 - Peso Normal";
        } else if (imc < 30) {
            ret = "< 30 - Acima do Peso";
        } else if (imc < 35) {
            ret = "< 35 - Obesidade I";
        } else if (imc < 40) {
            ret = "< 40 - Obesidade II";
        } else {
            ret = ">= 40 - Obesidade III";
        }
        return ret;
    }
    public void mensagem() {
        double imc = Math.Round(calculoIMC());
        string msg = retorno();
        Console.WriteLine("calculo: "+ imc +" = "+ peso +" / "+ altura +" * "+ altura);
        Console.WriteLine(msg);
    }
}
}