import os
from openai import OpenAI
import pandas as pd
from fuzzywuzzy import fuzz
import re
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Verify the API key
print(f"API Key: {os.environ.get('OPENAI_API_KEY')}")


# Initialize the OpenAI client
client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))

if not client.api_key:
    raise ValueError("The OPENAI_API_KEY environment variable is not set.")
else:
    print(f"API Key is set")

# Function to interact with OpenAI API and print raw response
def categorize_data(unstructured_text):
    # Examples to provide context about the data structure
    examples = """
    
    Example 1:
    Firm: Millennium Management
    Name: Sara Yong
    Title: Chief Operating Officer
    Region: APAC
    Location: Singapore
    Function: Executive Committee
    Strategy: Not specified
    Products: Not specified

    Example 2:
    Firm: Millennium Management
    Name: Michael Huttman
    Title: Chairman, Founder
    Region: Europe
    Location: London, UK
    Function: Executive Committee
    Strategy: Not specified
    Products: Equities

    Example 3:
    Firm: Millennium Management
    Name: Martin Pabari
    Title: Managing Director, CEO
    Region: Europe
    Location: London, UK
    Function: Executive Committee
    Strategy: Not specified
    Products: Not specified

    Example 4:
    Firm: Millennium Management
    Name: David Meneret
    Title: Founder, Chief Investment Officer
    Region: North America
    Location: New York, NY
    Function: Executive Committee
    Strategy: RV, Market Neutral
    Products: Derivatives, Not specified


        These are the total possible functions:

        Chief Investment Officer
        Executive Committee
        Senior Portfolio Manager
        Portfolio Manager
        Junior Portfolio Manager
        Portfolio Manager (Group Head)
        Portfolio Analyst
        Investor
        Investment Analyst
        Investment Analyst (Group Head)
        Proprietary Trader
        Trader
        Trader (Group Head)
        Quant Trader
        Quant Trader (Group Head)
        Execution Trader
        Execution Trader (Group Head)
        Trading Assistant
        Trading Assistant (Group Head)
        Repo Trader
        Repo Trader (Group Head)
        Quant Research
        Quant Research (Group Head)
        Quant Analyst
        Quant Analyst (Group Head)
        Strategist
        Strategist (Group Head)
        Economist
        Economist (Group Head)
        Business Development
        Business Development (Group Head)
        Investor Relations
        Investor Relations (Group Head)
        Sales
        Sales (Group Head)
        Business Manager
        Business Manager (Group Head)
        Capital Markets
        Capital Markets (Group Head)
        Portfolio Finance
        Portfolio Finance (Group Head)
        Product Control
        Product Control (Group Head)
        Treasury
        Treasury (Group Head)
        Risk Management
        Risk Management (Group Head)
        Quant Developer
        Quant Developer (Group Head)
        Developer
        Developer (Group Head)
        Data Science
        Data Science (Group Head)
        Valuation
        Valuation (Group Head)
        Technology
        Technology (Group Head)
        Middle Office
        Middle Office (Group Head)
        Operations
        Operations (Group Head)
        Compliance
        Compliance (Group Head)
        Counsel
        Product Specialist
        Back Office

        These are all possible regions:

        North America
        APAC
        MENA
        Europe
        LATAM

        These are the most common strategies:

        L/S Equities
        Credit
        Macro
        Equities
        Commodities
        Fixed Income
        Global Macro
        EM
        Rates
        Structured Products
        Fixed Income RV
        Structured Credit
        EM Macro
        Quant
        Systematic Macro
        Macro RV
        Systematic
        L/S Credit
        Equity Vol
        Macro Vol
        Convertible Arbitrage
        Rates Vol
        Equity Derivatives
        QIS
        Systematic Equities
        Private Credit
        Macro Credit
        Event-Driven
        Index Rebal
        Quant Macro
        Global Rates
        FX
        Global Macro Vol
        Macro Rates
        Private Equity
        Global Credit
        Index Arbitrage
        FX Vol
        Quant Equities
        Quantitative Strategies
        Global Equities
        Rates & Inflation
        Rates RV
        EM Rates & FX
        Global Fixed Income
        Discretionary Macro
        Global Discretionary Macro
        Global Macro RV
        Distressed Credit, Special Situations
        EM Credit
        Mortgage Strategies
        Opportunistic Credit
        Cross-Asset Vol
        Multi-Asset
        Stat Arb
        Global FX
        Global EM Macro
        European Rates
        Real Estate
        Fundamental Equities
        Macro Equity
        Vol
        Short Macro
        Municipal Bonds
        HFT
        EM FX
        Global EM
        Distressed Credit
        Capital Structure Arbitrage
        Agency MBS
        Systematic Credit
        Rates & FX
        Insurance
        Mortgages
        Cross-Asset
        Systematic Strategies
        Cross-Commodities
        TBD
        Vol RV
        Systematic Fixed Income
        Real Assets
        Venture Capital
        L/S
        Algorithmic Trading
        MBS
        Equity Vol RV
        Global Systematic Macro
        Macro Vol RV
        Credit, Equities
        Cross-Asset Macro
        Algo Trading
        Event-Driven, Merger Arbitrage
        Global Rates & FX
        Short Macro / STIR
        Global L/S Equities
        Bond RV
        Global Convertible Arbitrage
        Asia Macro
        Macro FX
        Multi-Strat
        Cross-Asset Macro Vol
        Global Rates & Inflation
        Global Rates RV
        Systematic Commodities
        European Rates RV
        Crypto
        Value Equities
        Merger-Arbitrage, Event-Driven
        Macro Commodities
        Special Situations
        Equity Index Vol
        Leveraged Finance
        Asset Allocation
        European Equities
        Liquid Rates
        Event-Driven Credit

    Please categorize the following text based on the examples above:
    """

    # Combine examples with the unstructured text
    prompt = examples + "\n" + unstructured_text

    response = client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[
            {
                "role": "user",
                "content": f"{prompt}",
            }
        ]
    )
    # Print the raw response (Remove # to print raw response)
    #print("Raw response from API:")
    #print(response.model_dump_json(indent=2))
    
    # Extract the message content
    extracted_text = response.choices[0].message.content.strip()
    return extracted_text

# Function to parse the extracted text
def parse_extracted_text(extracted_text):
    categories = ["Firm", "Name", "Title", "Region", "Location", "Function", "Strategy", "Products"]
    data = {}
    for category in categories:
        data[category] = ""
        for line in extracted_text.split("\n"):
            if line.startswith(category):
                data[category] = line.split(":")[1].strip()
    return data

# Function to test the categorize_data function in the terminal
def test_categorize_data():
    while True:
        prompt = input("Enter the text to categorize (type 'END' on a new line to finish): ")
        lines = []
        while True:
            line = input()
            if line.strip().upper() == 'END':
                break
            lines.append(line)
        prompt = "\n".join(lines)
        if prompt.lower() == 'exit':
            break
        extracted_text = categorize_data(prompt)
        parsed_data = parse_extracted_text(extracted_text)
        print("Extracted Data:")
        for key, value in parsed_data.items():
            print(f"{key}: {value}")

if __name__ == "__main__":
    test_categorize_data()
