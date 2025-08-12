import styled from 'styled-components';

export const StyledContainer = styled.div`
  text-align: center;
  padding: 15px;
  background-color: white;
  font-family: Arial, sans-serif;
`;

export const StyledTitle = styled.h1`
  color: #217346;
  font-size: 24px;
  margin-bottom: 10px;
`;

export const StyledText = styled.p`
  font-size: 14px;
  line-height: 1.4;
  margin-bottom: 15px;
`;

export const StyledButton = styled.button`
  background-color: #217346;
  color: white;
  border: none;
  padding: 8px 15px;
  font-size: 16px;
  cursor: pointer;
  border-radius: 5px;
  margin-top: 10px;

  &:hover {
    background-color: #1a5c38;
  }
`;

export const StyledResetButton = styled(StyledButton)`
  background-color: #f44336; /* Red color for reset */

  &:hover {
    background-color: #d32f2f;
  }
`;

export const ButtonContainer = styled.div`
  margin-top: 8px;
  display: flex;
  justify-content: center;
  gap: 10px;
`;
