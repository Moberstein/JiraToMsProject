namespace JiraToMsProject
{
    class Issue
    {
        public string Name { get; set; }
        public string Ressource { get; set; }
        public string[] Sprints { get; set; }
        public string Key { get; set; }
        public string Type { get; set; }
        public string Estimated { get; set; }
        public string SubTasks { get; set; }
        public string Linked { get; set; }
        public string[] Labels { get; set; }
        public object Epic { get; set; }
        public string Created { get; set; }
    }
}
