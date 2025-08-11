import styled from "styled-components";

export const StyledButton = styled.button`
  background-color: #f44336; /* Red color for reset */
  color: white;
  border: none;
  padding: 8px 15px; /* Reduced padding */
  font-size: 16px; /* Slightly smaller font size */
  cursor: pointer;
  border-radius: 5px;
  margin-top: 15px; /* Reduced margin */

  &:hover {
    background-color: #d32f2f;
  }
`;

export const StyledNavButton = styled(StyledButton)`
  background-color: #217346;
  margin: 3px; /* Reduced margin */

  &:hover {
    background-color: #1a5c38;
  }

  &:disabled {
    background-color: #cccccc;
    cursor: not-allowed;
  }
`;

export const StyledFooter = styled.div`
  position: fixed;
  bottom: 0;
  left: 0;
  width: 100%; /* Set width to 100% */
  box-sizing: border-box; /* Include padding in width calculation */
  background-color: #f0f0f0; /* Light gray background for footer */
  padding: 10px 15px;
  display: flex;
  justify-content: space-between; /* Distribute items with space between them */
  align-items: center;
  box-shadow: 0 -2px 5px rgba(0, 0, 0, 0.1);
  z-index: 1000;
`;

export const StyledNavButtonsContainer = styled.div`
  display: flex;
  gap: 10px; /* Space between nav buttons */
  align-items: center;
`;
