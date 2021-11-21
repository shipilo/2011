namespace Les_2011
{
    internal class Student
    {
        public string Name;
        public int Group;
        public double Ratio;

        public Student(string name, int group)
        {
            Name = name;
            Group = group;
            Ratio = 1;
        }

        public override string ToString()
        {
            return $"Имя: {Name} - Группа: {Group}";
        }
    }
}