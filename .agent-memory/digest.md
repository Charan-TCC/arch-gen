# Project Digest — Arch Gen

Re-engineered the specification from a simple linear flow to a highly-available, production-grade Multi-AZ architecture layout.
- Added explicit structural boundaries: AWS Cloud -> Production VPC -> AZ A / AZ B.
- Split EKS, RDS, and ElastiCache nodes across AZ A (primary) and AZ B (standby).
- Recreated the edge tier (Route53, CloudFront, WAF) properly routing to a centralized ALB.
- Output PNG updated successfully.
