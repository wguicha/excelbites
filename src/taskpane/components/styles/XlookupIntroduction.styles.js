import styled from "styled-components";

export const StyledContainer = styled.div`
  text-align: center;
  padding: 10px; /* Further reduced padding */
  background-color: white;
  font-family: Arial, sans-serif;
`;

export const StyledLogo = styled.img`
  max-width: 100px; /* Even smaller logo */
  margin-bottom: 10px; /* Reduced margin */
`;

export const StyledTitle = styled.h1`
  color: #217346;
  font-size: 22px; /* Further smaller font size */
  margin-bottom: 8px; /* Reduced margin */
`;

export const StyledParagraph = styled.p`
  font-size: 13px; /* Further smaller font size */
  line-height: 1.3;
  margin-bottom: 10px; /* Reduced margin */
`;

export const StyledAdvantagesContainer = styled.div`
  margin: 10px 0; /* Reduced margin */
  padding: 8px; /* Reduced padding */
  border: 1px solid #e0e0e0;
  border-radius: 8px;
  background-color: #f9f9f9;
  text-align: left;
`;

export const StyledAdvantagesTitle = styled.h2`
  color: #217346;
  font-size: 16px; /* Further smaller font size */
  margin-bottom: 6px; /* Reduced margin */
  text-align: center;
`;

export const StyledAdvantagesList = styled.ul`
  list-style: none;
  padding: 0;
  margin: 0;
`;

export const StyledAdvantageItem = styled.li`
  font-size: 13px; /* Further smaller font size */
  margin-bottom: 4px; /* Reduced margin */
  display: flex;
  align-items: center;
`;

export const CheckMark = styled.span`
  color: #217346;
  font-size: 16px; /* Further smaller font size */
  margin-right: 6px; /* Reduced margin */
`;

export const StyledButton = styled.button`
  background-color: #217346;
  color: white;
  border: none;
  padding: 6px 12px; /* Further reduced padding */
  font-size: 14px; /* Further smaller font size */
  cursor: pointer;
  border-radius: 5px;
  margin: 2px; /* Reduced margin */
  min-width: 150px; /* Added min-width for consistent sizing */

  &:hover {
    background-color: #1a5c38;
  }
`;

export const StyledNavButton = styled(StyledButton)`
  background-color: #a9a9a9;

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
  margin-top: 6px; /* Reduced margin */
  display: flex;
  justify-content: center;
  gap: 8px; /* Reduced space between buttons */
`;
