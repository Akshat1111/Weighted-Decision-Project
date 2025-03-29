# Weighted-Decision-Project
# **Missing Data Imputation Algorithm**

## **Algorithm Overview**
Since my goal is to fill all missing values while maintaining data consistency, I needed to develop an algorithm that could be systematically applied across all cells with missing data.

To address this, I designed a decision-making algorithm that assigns weighted averages based on three key factors: **year, region, and state**. Given that the dataset is predominantly structured by years, I assigned:
- **50% weight to the yearly average**
- **30% to the regional average**
- **20% to the state average**  

However, the weights assigned to states and regions need to be flexible. Some states appear only once or twice in the dataset, and assigning them a fixed 20% weight could introduce inaccuracies, increasing the deviation from actual values.

### **Algorithm Logic**
To handle this, the algorithm follows these steps:
- If a cell has a missing value for a given year (e.g., `2004`), the predicted value is calculated as:  
0.5 * (Average of 2004) + 0.3 * (Average of Region) + 0.2 * (Average of State) (if state occurrence > 4)
- Otherwise, if the state appears **4 times or less**, its weight is reduced to `0`, and the region compensates:
0.5 * (Average of 2004) + 0.5 * (Average of Region) + 0 * (Average of State)
- To further refine this approach, I plan to implement a **dynamic weighting system** for states based on their occurrence frequency, adjusting the region’s weight accordingly.  
For example:
- If a state appears between `1 and 2 times`, `2 and 3 times`, etc., the state's contribution will be progressively reduced while ensuring a balanced weight distribution.

## **Implementation**
Initially, I attempted to implement this in Excel but, due to my limited knowledge of advanced Excel formulas, I faced challenges at the start. However, since I am already familiar with **Python and its data structures**, I developed a Python script that implements the algorithm efficiently.

### **Tools Used**
- **Python**
- **openpyxl** (to manipulate Excel files programmatically)

I have also uploaded the **final filled Excel file** in the repository. The code includes **comments** at necessary points to explain the logic clearly.

---

## **⚠️ Important Warning**
- The provided **Python scripts are specifically designed to work with the given Excel files**.
- If you want to test them on your **own dataset**, you **must modify the row and column indices** accordingly to match your worksheet structure.
- Failure to do so may lead to incorrect results or errors.

---

Feel free to explore the repository, test the code, and provide feedback! 🚀
