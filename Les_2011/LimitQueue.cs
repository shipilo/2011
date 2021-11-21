using System.Collections;

namespace Les_2011
{
    internal class LimitQueue<T> : Queue
    {
        public override void Enqueue(object obj)
        {
            base.Enqueue(obj);
            if (Count == 4)
            {
                if (obj is Draw)
                {
                    foreach (int index in (obj as Draw).Winners)
                    {
                        Program.data[index].Ratio = 1;
                    }
                }
                Dequeue();
            }
        }
    }
}
