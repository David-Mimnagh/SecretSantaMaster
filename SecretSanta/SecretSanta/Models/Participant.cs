using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SecretSanta.Models
{
    public class Participant
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public int SpouseId { get; set; }
        public string EmailAddress { get; set; }
        public string Interests { get; set; }
        public int PreviousSSantaId { get; set; }
        public string ShirtSize { get; set; }
    }
}
