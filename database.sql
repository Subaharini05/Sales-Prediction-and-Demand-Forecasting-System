-- Database Schema for Coconut Coir Sales Prediction and Demand Forecasting System

CREATE TABLE IF NOT EXISTS Users (
    UserID INT AUTO_INCREMENT PRIMARY KEY,
    Username VARCHAR(100) UNIQUE NOT NULL,
    Password VARCHAR(255) NOT NULL,
    Role VARCHAR(50) NOT NULL
);

CREATE TABLE IF NOT EXISTS Products (
    ProductID INT AUTO_INCREMENT PRIMARY KEY,
    ProductName VARCHAR(255) NOT NULL,
    Category VARCHAR(100) NOT NULL,
    Price DECIMAL(10, 2) NOT NULL
);

CREATE TABLE IF NOT EXISTS Sales (
    SaleID INT AUTO_INCREMENT PRIMARY KEY,
    Date DATE NOT NULL,
    ProductID INT NOT NULL,
    Quantity INT NOT NULL,
    Price DECIMAL(10, 2) NOT NULL,
    Revenue DECIMAL(10, 2) NOT NULL
);

CREATE TABLE IF NOT EXISTS Predictions (
    PredictionID INT AUTO_INCREMENT PRIMARY KEY,
    ProductName VARCHAR(255) NOT NULL,
    PredictionDate DATE NOT NULL,
    PredictedSales DOUBLE NOT NULL
);

CREATE TABLE IF NOT EXISTS DemandForecast (
    ForecastID INT AUTO_INCREMENT PRIMARY KEY,
    ProductName VARCHAR(255) NOT NULL,
    ForecastQuantity DOUBLE NOT NULL,
    DemandLevel VARCHAR(50) NOT NULL
);
