# Library-Orders-Processor

Mover Orders though a simple process of 'To Be Ordered' -> 'On Order' -> 'Received'
Each Department can have it's own flow so they know the status of thier own books

To move data though the flow sheets must be named in the following format:
 * 'To Be Ordered - YOUR_NAME_HERE'
 * 'On Order - YOUR_NAME_HERE'
 * 'Received'

To move an order each sheet expects a specific status to be entered in 'Order Status' feild
Entering the correct status causes the order to be moved to the next step
 * To Be Ordered -> Ordered
 * On Order -> Received
