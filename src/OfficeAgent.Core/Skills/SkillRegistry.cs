using System;
using System.Collections.Generic;
using OfficeAgent.Core.Services;

namespace OfficeAgent.Core.Skills
{
    public sealed class SkillRegistry : ISkillRegistry
    {
        private readonly Dictionary<string, IAgentSkill> skills;

        public SkillRegistry(params IAgentSkill[] skills)
        {
            this.skills = new Dictionary<string, IAgentSkill>(StringComparer.OrdinalIgnoreCase);
            foreach (var skill in skills ?? Array.Empty<IAgentSkill>())
            {
                if (skill == null)
                {
                    continue;
                }

                this.skills[skill.SkillName] = skill;
            }
        }

        public IAgentSkill Resolve(string skillName)
        {
            if (string.IsNullOrWhiteSpace(skillName))
            {
                throw new ArgumentException("Skill resolution requires a skill name.", nameof(skillName));
            }

            if (!skills.TryGetValue(skillName.Trim(), out var skill))
            {
                throw new ArgumentException($"Skill '{skillName}' is not registered.", nameof(skillName));
            }

            return skill;
        }
    }
}
