﻿namespace MothersonAuditCheckList.Models.DTO
{
    public class AuditListDTO
    {
        public string Unit { get; set; }
        public DateTime AuditDate { get; set; } 
        public string auditees { get; set; }
        public List<Rule> RuleList {  get; set; }
    }
    public class Rule
    {
        public string RuleName { get; set; }
        public List <RuleDetail> RuleListDetails { get; set; }
    }

    public class RuleDetail
    {
        public string Section { get; set; }
        public string Type { get; set; }
        public string Statement { get; set; }
        public string Score { get; set; }
        public string Remark { get; set; }
    }
}