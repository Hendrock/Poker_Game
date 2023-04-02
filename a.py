import random
from openpyxl import Workbook

# Create a deck of cards
suits = ['Hearts', 'Diamonds', 'Clubs', 'Spades']
ranks = {str(i): i for i in range(2, 11)}
ranks.update({'Jack': 11, 'Queen': 12, 'King': 13, 'Ace': 14})

# Function to check for hand types
def check_hand(hand):
    hand_rank = [card[0] for card in hand]
    hand_suit = [card[1] for card in hand]
    num_aces = hand_rank.count('Ace')
    if len(set(hand_suit)) == 1:
        if set(hand_rank) == {10, 11, 12, 13, 14}:
            return 'ROYAL FLUSH', num_aces
        elif hasStraight([ranks[rank] for rank in hand_rank]):
            return 'STRAIGHT FLUSH', num_aces
        else:
            return 'FLUSH', num_aces
    elif hasStraight([ranks[rank] for rank in hand_rank]):
        return 'STRAIGHT', num_aces
    elif len(set(hand_rank)) == 2:
        if hand_rank.count(hand_rank[0]) == 4 or hand_rank.count(hand_rank[1]) == 4:
            return '4 OF A KIND', num_aces
        else:
            return 'FULL HOUSE', num_aces
    elif len(set(hand_rank)) == 3:
        if hand_rank.count(hand_rank[0]) == 3 or hand_rank.count(hand_rank[1]) == 3 or hand_rank.count(hand_rank[2]) == 3:
            return '3 OF A KIND', num_aces
        else:
            return '2 PAIR', num_aces
    elif len(set(hand_rank)) == 4:
        return '1 PAIR', num_aces
    else:
        return 'HIGH CARD', num_aces
    
# Function to check for a straight
def hasStraight(hand):
    # create a set of the rank values for the hand, with low-ace converted to 1
    ranks_set = set([(1 if rank == 'Ace' else int(rank)) for rank, suit in hand])
    if len(ranks_set) < 5:
        return False
    # check for a straight using both low-ace and high-ace values
    ranks_high = sorted([14 if rank == 1 else int(rank) for rank, suit in hand])
    ranks_low = sorted([1 if rank == 'Ace' else int(rank) for rank, suit in hand])
    if ranks_high[-1] - ranks_high[0] == 4 or (ranks_high[-1] == 14 and ranks_high[-2] == 5 and ranks_high[0] == 2):
        return True
    elif ranks_low[-1] - ranks_low[0] == 4:
        return True
    else:
        return False

# Function to deal cards
def deal_cards():
    deck = [(rank, suit) for rank in ranks for suit in suits]
    hand = []
    for i in range(5):
        card = random.choice(deck)
        hand.append(card)
    return hand


# Main function
def main():
    wb = Workbook()
    ws = wb.active
    ws.title = 'Poker Hands'
    headers = ['Hand', 'Hand Type', 'Number of Ace']
    ws.append(headers)
    
    for i in range(50000):
        hand = deal_cards()
        hand_str = ', '.join([f'{card[0]} of {card[1]}' for card in hand])
        hand_type, num_aces = check_hand(hand)
        row = [hand_str, hand_type, num_aces]
        ws.append(row)
        
    wb.save('poker_handsfinalfinalfinalfinal.xlsx')

if __name__ == '__main__':
    main()
