import styled from "styled-components";

export const StyledContainer = styled.div`
  text-align: center;
  padding: 15px; /* Reduced padding */
  background-color: white;
  font-family: Arial, sans-serif;
`;

export const StyledTitle = styled.h1`
  color: #217346;
  font-size: 24px; /* Slightly smaller font size */
  margin-bottom: 10px; /* Reduced margin */
`;

export const StyledText = styled.p`
  font-size: 14px; /* Slightly smaller font size */
  line-height: 1.4;
  margin-bottom: 15px; /* Reduced margin */
`;

export const StyledForm = styled.div`
  display: flex;
  flex-direction: column;
  align-items: flex-start;
  margin: 0 auto;
  max-width: 280px; /* Slightly reduced max-width */
  padding: 15px; /* Reduced padding */
  border: none;
  border-radius: 0;
  background-color: white;
  box-shadow: none;
`;

export const StyledLabel = styled.label`
  margin-top: 8px; /* Reduced margin */
  font-weight: bold;
  text-align: left;
  width: 100%;
  font-size: 14px; /* Slightly smaller font size */
`;

export const StyledInput = styled.input`
  width: 100%;
  padding: 6px; /* Reduced padding */
  margin-top: 3px; /* Reduced margin */
  border: 1px solid #ddd;
  border-radius: 4px;
  font-size: 14px; /* Slightly smaller font size */
`;

export const StyledButton = styled.button`
  background-color: #217346;
  color: white;
  border: none;
  padding: 6px 12px; /* Further reduced padding */
  font-size: 14px; /* Further smaller font size */
  cursor: pointer;
  border-radius: 5px;
  margin-top: 10px; /* Reduced margin */
  min-width: 150px; /* Added min-width for consistent sizing */

  &:hover {
    background-color: #1a5c38;
  }

  &:disabled {
    background-color: #cccccc;
    cursor: not-allowed;
  }
`;

export const StyledNavButton = styled(StyledButton)`
  background-color: #a9a9a9;
  margin: 3px; /* Reduced margin */

  &:hover {
    background-color: #808080;
  }
`;

export const StyledResetButton = styled(StyledButton)`
  background-color: #f44336; /* Red color for reset */

  &:hover {
    background-color: #d32f2f;
  }
`;

export const ButtonContainer = styled.div`
  margin-top: 8px; /* Reduced margin */
  display: flex;
  justify-content: center;
  gap: 10px; /* Space between buttons */
`;

export const StyledMessage = styled.p`
  color: #217346;
  font-weight: bold;
  margin-top: 8px; /* Reduced margin */
  background-color: #e6ffe6;
  border: 1px solid #217346;
  padding: 4px; /* Reduced padding */
  border-radius: 4px;
  font-size: 14px; /* Slightly smaller font size */
`;
