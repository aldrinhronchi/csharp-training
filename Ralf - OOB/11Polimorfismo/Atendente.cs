using System;

namespace _11Polimorfismo
{
    class Atendente : Imposto
    {
        public override void valeAlimentacao(double salario)
        {
            Console.WriteLine("Desconto Atendente do vale alimentação R$" + (salario * 0.12));
        }
    }
}