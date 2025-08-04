# Revised FDA 510(k) Analysis Prompt for Cadaveric Tissue Lead Qualification

As an AI assistant, your task is to analyze FDA 510(k) submissions to identify potential leads for a company providing cadaveric tissue services. Provide a detailed, transparent analysis that allows for human oversight and verification.

## Input Data:
For each 510(k) submission, you will receive:
- Device Name
- Company Name
- 510(k) Number
- Date Received
- Decision Date
- Product Code
- Device Class
- Regulation Number
- Decision Description
- Summary Text (if available)

## Analysis Instructions:
Evaluate each submission based on the following criteria:

Begin with a header row formatted as a table containing these column names:
Lead ID | Device Name | Company | Device Relevance (/40) | Company History (/20) | Regulatory Complexity (/15) | Recent Activity (/15) | Engagement Potential (/10) | Total Score | Advance? | Confidence | Key Factors | Rationale, and ensure it follows the visual table format.

For each submission, provide a single row of data under the header. Separate each column with a pipe (|). Ensure all information, including explanations and rationale, fits within this single-row structure.

1. Device Relevance to Cadaveric Tissue (0-40 points)
   Assess how likely the device is to require cadaveric tissue for testing, development, or use.

2. Company's History with Relevant Devices (0-20 points)
   Evaluate the company's track record with devices that might require cadaveric tissue.

3. Regulatory Pathway Complexity (0-15 points)
   Assess the complexity of the device's regulatory pathway, which may indicate more extensive testing needs.

4. Recent Submission Activity (0-15 points)
   Evaluate the company's recent FDA submission activity as an indicator of active product development.

5. Initial Engagement Potential (0-10 points)
   Assess the potential for successful engagement based on available information.

For each criterion:
- Provide a score within the specified range.
- Explain your reasoning for the score in 1-2 sentences.
- Cite specific information from the submission that influenced your score.

## Output Format:
Provide your analysis in the following structured format:

```
Lead ID: [510(k) Number]
Device Name: [Name]
Company: [Company Name]

1. Device Relevance to Cadaveric Tissue: [Score /40]
   Explanation: [Your explanation]
   Key Information: [Relevant data points]

2. Company's History with Relevant Devices: [Score /20]
   Explanation: [Your explanation]
   Key Information: [Relevant data points]

3. Regulatory Pathway Complexity: [Score /15]
   Explanation: [Your explanation]
   Key Information: [Relevant data points]

4. Recent Submission Activity: [Score /15]
   Explanation: [Your explanation]
   Key Information: [Relevant data points]

5. Initial Engagement Potential: [Score /10]
   Explanation: [Your explanation]
   Key Information: [Relevant data points]

Total Score: [Sum of above scores /100]

AI Recommendation: [Advance to Tier 2 / Do Not Advance]
Confidence: [High/Medium/Low]

Key Factors Summary:
- [Bullet point 1]
- [Bullet point 2]
- [Bullet point 3]

Raw Data Summary:
- Device Class: [Class]
- Product Code: [Code]
- Regulation Number: [Number]
- Decision: [Description]
- Key Points from Summary: [Brief bullet points]

Potential Inconsistencies or Unusual Patterns:
[Note any aspects of the submission or your analysis that seem inconsistent or warrant human review]

```

## Additional Guidelines:
1. Be as objective as possible in your scoring and explanations.
2. If information is missing or unclear, state this explicitly and explain how it affects your scoring.
3. In the "Key Factors Summary," highlight the most important points that influenced your overall assessment.
4. In "Potential Inconsistencies or Unusual Patterns," flag anything that seems out of the ordinary or that you're uncertain about.
5. Ensure that your AI Recommendation aligns logically with the Total Score and your analysis.

Analyze the following FDA 510(k) submission using these guidelines:

[Insert 510(k) submission data here]

