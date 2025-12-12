# app/core/validator.py
class Validator:
    def __init__(self):
        pass
    
    def validate_time_format(self, time_str):
        """Validasi format waktu"""
        import re
        pattern = r'^\d{1,2}[:\.]\d{2}\s*[-–]\s*\d{1,2}[:\.]\d{2}$'
        return bool(re.match(pattern, str(time_str)))
    
    def validate_dataframe(self, df, required_columns):
        """Validasi dataframe"""
        errors = []
        
        # Check required columns
        missing_cols = [col for col in required_columns if col not in df.columns]
        if missing_cols:
            errors.append(f"Missing columns: {missing_cols}")
        
        # Check empty dataframe
        if df.empty:
            errors.append("Dataframe is empty")
        
        return errors
