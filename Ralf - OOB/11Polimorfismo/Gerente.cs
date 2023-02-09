using System;

namespace _11Polimorfismo
{
    class Gerente : Imposto
    {
        //Metodo 
        public override void valeAlimentacao(double salario)
        {
            Console.WriteLine("Desconto Gerente do vale alimentação R$" + (salario * 0.15));
        }
    }
}